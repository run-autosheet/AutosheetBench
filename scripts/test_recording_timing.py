"""No UI/pilot execution: test recording math, lifecycle isolation and write guards."""
import hashlib
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parent))
import recording_timing as timing
from run_excel import PublicRecordingRunner, load_engine, main
from claude_transcript import visible_reply

def video(directory):
    (directory / "recording.mp4").write_bytes(b"fake video; unit tests do not record a product")
    return {"status": "success", "file": "recording.mp4", "timingSchemaVersion": "atomic-recording-clock/v1",
            "firstFrameEpochMs": 1_000_000, "clockMappingUncertaintyMs": 0.5,
            "durationSeconds": 201, "frameOffsetsSeconds": list(range(201)),
            "framesPerSecond": 5, "timeCompressionApplied": False}

def observed_turns():
    return [{"status": "ok", "submittedAtEpochMs": 1_010_000, "timingEndpointEpochMs": 1_050_000},
            {"status": "ok", "submittedAtEpochMs": 1_090_000, "timingEndpointEpochMs": 1_150_000}]

class RecordingMathTests(unittest.TestCase):
    def test_each_turn_uses_original_frame_clock_and_excludes_the_gap(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            recording = video(root)
            payload = timing.timing_candidates("sample", recording, observed_turns(), root, "success")
            self.assertEqual(payload["candidateDurationSeconds"], 100)
            self.assertEqual(payload["turns"][1]["startSeconds"], 90)

    def test_recording_is_required_and_badges_cannot_fill_a_missing_endpoint(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            recording = video(root)
            for broken in [{**recording, "status": "error"}, {**recording, "timeCompressionApplied": True},
                           {**recording, "firstFrameEpochMs": None}, {**recording, "frameOffsetsSeconds": [0, 1, 1]}]:
                with self.subTest(broken=broken):
                    result = timing.timing_candidates("sample", broken, observed_turns(), root, "success")
                    self.assertEqual(result["status"], "unavailable")
            turns = observed_turns()
            turns[0].pop("timingEndpointEpochMs")
            turns[0]["productReportedDurationSeconds"] = 3
            self.assertEqual(timing.timing_candidates("sample", recording, turns, root, "success")["status"], "unavailable")
            recovered = observed_turns()
            recovered[-1]["status"] = "error"
            self.assertEqual(timing.timing_candidates("sample", recording, recovered, root, "success")["status"], "unavailable")

class PublicLifecycleTests(unittest.TestCase):
    def test_chat_fallback_rejects_mixed_accessibility_text_and_prompt_echo(self):
        self.assertIsNone(visible_reply({"status": "ok", "source": "accessibility-tree", "lines": ["1 text Done"]}, "prompt"))
        self.assertIsNone(visible_reply({"status": "ok", "source": "claude-copy-message", "text": "prompt"}, "prompt"))
        self.assertEqual(visible_reply({"status": "ok", "source": "claude-copy-message", "text": "Actual answer"}, "prompt"), "Actual answer")

    def test_chat_fallback_is_turn_complete_and_never_claims_a_native_export(self):
        engine = load_engine(); runner = PublicRecordingRunner(engine)
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary); task = SimpleNamespace(task_id="sample", prompts=("one", "two"))
            runner.turns = [{"status": "ok"}, {"status": "ok"}]
            runner.chat_turns = [{"turn": 1, "text": "First", "captureSha256": "a"}, {"turn": 2, "text": "Second", "captureSha256": "b"}]
            native = {"status": "error", "error": "native format unavailable"}
            self.assertTrue(runner.write_transcript(task, root, "claude", "model", "high", {}, native))
            data = timing.read_json(root / "transcript.json")
            self.assertFalse(data["nativeTranscriptCaptured"])
            self.assertEqual([e["text"] for e in data["entries"]], ["one", "First", "two", "Second"])
            runner.original["write_transcript"] = Mock(return_value=False)
            for mutation in (lambda: runner.chat_turns.pop(), lambda: runner.turns[-1].update(status="timeout"), lambda: native.update(inputRequestPending=True)):
                mutation()
                self.assertFalse(runner.write_transcript(task, root, "claude", "model", "high", {}, native))
                self.assertFalse(runner.chat_fallback_accepted)

    def test_followup_cannot_submit_when_fresh_layout_verification_fails(self):
        engine = load_engine(); runner = PublicRecordingRunner(engine)
        temporary = tempfile.TemporaryDirectory(); self.addCleanup(temporary.cleanup)
        runner.directory = Path(temporary.name); runner.turns = [{"status": "ok"}]; runner.recording = object()
        runner.args = SimpleNamespace(model="model", reasoning="high", skip_config_validation=False)
        runner.original["osascript_op"] = Mock(return_value={"status": "ok", "sendButtonCenter": [1, 2]})
        with patch.object(engine, "verify_recording_before_submission", return_value={"status": "error", "error": "wrong target"}):
            result = runner.operation("send", product="claude", **{"--xlsx": Path("task.xlsx"), "--prompt-file": Path("prompt.txt")})
        self.assertEqual(result["status"], "error")
        self.assertEqual(runner.original["osascript_op"].call_count, 1)
        self.assertEqual(runner.original["osascript_op"].call_args.kwargs["--stage-only"], "true")

    def test_layout_history_preserves_both_turns(self):
        runner = PublicRecordingRunner(load_engine())
        with tempfile.TemporaryDirectory() as temporary:
            runner.directory = Path(temporary)
            def layout(*a, **k):
                (runner.directory / "recording-layout.png").write_bytes(str(len(runner.turns)).encode())
                timing.write_json(runner.directory / "recording-layout.json", {"turn": len(runner.turns) + 1})
                return {"status": "ok"}
            runner.original["verify_recording_before_submission"] = layout
            runner.verify_recording_before_submission(None, None); runner.turns.append({"status": "ok"})
            runner.verify_recording_before_submission(None, None)
            self.assertEqual((runner.directory / "recording-layout-turn-01.png").read_bytes(), b"0")
            self.assertEqual((runner.directory / "recording-layout-turn-02.png").read_bytes(), b"1")

    def test_runner_and_node_capture_share_the_native_session_lease(self):
        runner = PublicRecordingRunner(load_engine())
        other = PublicRecordingRunner(load_engine())
        with tempfile.TemporaryDirectory() as temporary, patch.dict(os.environ, {"AUTOSHEETBENCH_NATIVE_LOCK_DIR": temporary}):
            previous = os.environ.get("AUTOSHEETBENCH_EXCEL_SESSION_TOKEN")
            runner.acquire_session_lock()
            lock = runner.session_lock
            try:
                with self.assertRaises(FileExistsError):
                    other.acquire_session_lock()
                script = """
import {acquireCaptureLocks} from './scripts/capture_locks.mjs';
const release=acquireCaptureLocks({taskDir:process.argv[1],platform:'microsoft_excel'});
release();
try {
  acquireCaptureLocks({taskDir:process.argv[1],platform:'microsoft_excel',retryMissing:true,captureRole:'evaluator'});
  process.exitCode=1;
} catch(error) {
  if(!error.message.includes('session is busy')) throw error;
  console.log('runner-owned capture allowed; evaluator retry blocked');
}
"""
                result = subprocess.run(["node", "--input-type=module", "-e", script, temporary],
                                        cwd=Path(__file__).resolve().parents[1], capture_output=True, text=True, check=True)
                self.assertIn("evaluator retry blocked", result.stdout)
                self.assertTrue(lock.is_file())
            finally:
                runner.release_lock()
            self.assertFalse(lock.exists())
            self.assertEqual(os.environ.get("AUTOSHEETBENCH_EXCEL_SESSION_TOKEN"), previous)

    def test_install_is_isolated_and_engine_duration_policy_is_unchanged(self):
        engine, public = load_engine(), load_engine()
        original_run_task = engine.run_task
        runner = PublicRecordingRunner(public)
        runner.install()
        self.assertIs(engine.run_task, original_run_task)
        self.assertNotEqual(public.RECORDING_BINARY, engine.RECORDING_BINARY)
        self.assertEqual(public.CAPTURE_SCRIPT.name, 'capture_workbook.mjs')
        self.assertEqual(engine.CAPTURE_SCRIPT.name, 'capture_run_workbook.mjs')
        self.assertEqual(engine.primary_duration_seconds(100, {"completionSeconds": 140},
                         product_reported_turn_durations=[3, 4])[0], 7)
        self.assertTrue(public.external_config_id("gptforwork", "test", "high").endswith("--" + timing.POLICY))

    def test_old_macos_is_refused_before_any_product_execution(self):
        with patch("run_excel.platform.mac_ver", return_value=("14.7", (), "")):
            with self.assertRaisesRegex(SystemExit, "macOS 15"):
                main(["--task-ids", "add_days"])

    def test_recording_watchdog_covers_excluded_gaps_without_changing_execution_budget(self):
        engine = load_engine()
        runner = PublicRecordingRunner(engine)
        runner.args = SimpleNamespace(send_timeout=90, transcript_timeout=60)
        task = engine.TaskSpec("sample", Path("."), {}, Path("fake.xlsx"), ("one", "two"), ("a", "b"), 300)
        capture = Mock(return_value=(None, None))
        runner.original["start_task_recording"] = capture
        runner.start_task_recording(task=task)
        self.assertEqual(task.time_limit_seconds, 300)
        self.assertEqual(capture.call_args.kwargs["task"].time_limit_seconds, 630)

    def test_terminal_suffix_excludes_confirmation_delay_but_not_permission_waits(self):
        engine = load_engine()
        runner = PublicRecordingRunner(engine)
        observations = iter([
            {"status": "ok", "state": "busy", "observedAtEpochMs": 1_020_000},
            {"status": "ok", "state": "idle", "observedAtEpochMs": 1_040_000, "permission": {"button": "Allow"}},
            {"status": "ok", "state": "idle", "observedAtEpochMs": 1_050_000},
            {"status": "ok", "state": "idle", "observedAtEpochMs": 1_054_000},
        ])
        runner.original["osascript_op"] = lambda *a, **k: next(observations)
        def wait(*args, **kwargs):
            for _ in range(4):
                runner.operation("check")
            return {"status": "ok", "sawBusy": True, "lastObservedAtEpochMs": 1_054_000, "elapsedSeconds": 44}
        runner.original["wait_for_turn"] = wait
        result = runner.wait_for_turn(Path("fake.xlsx"), submitted_at_epoch_ms=1_010_000,
                                      deadline=time_deadline(engine, 300))
        self.assertEqual(result["elapsedSeconds"], 40)
        self.assertEqual(result["timingEndpointEpochMs"], 1_050_000)

    def test_timeout_uses_the_actual_budget_cutoff_even_if_observation_is_stale(self):
        engine = load_engine()
        runner = PublicRecordingRunner(engine)
        runner.original["wait_for_turn"] = lambda *a, **k: {
            "status": "timeout", "lastObservedAtEpochMs": 1_020_000, "elapsedSeconds": 10}
        with patch.object(engine, "monotonic_from_epoch_ms", return_value=500):
            result = runner.wait_for_turn(Path("fake.xlsx"), submitted_at_epoch_ms=1_010_000, deadline=560)
        self.assertEqual(result["elapsedSeconds"], 60)
        self.assertEqual(result["timingEndpointEpochMs"], 1_070_000)

    def test_public_metadata_has_no_unverified_primary_duration_or_badge_fallback(self):
        engine = load_engine()
        runner = PublicRecordingRunner(engine)
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            recording = video(root)
            runner.turns = observed_turns()
            task = SimpleNamespace(task_id="sample", prompts=("one", "two"))
            def engine_write(**kwargs):
                timing.write_json(kwargs["task_output_dir"] / "metadata.json", {
                    "taskId": "sample", "status": "success", "durationSeconds": 7,
                    "durationMs": 7000, "finishedInTime": True,
                    "productReportedDurationSeconds": 7, "runner": {"screenRecording": recording}})
            runner.original["write_metadata"] = engine_write
            runner.write_metadata(task_output_dir=root, task=task, screen_recording=recording)
            metadata = timing.read_json(root / "metadata.json")
            self.assertIsNone(metadata["durationSeconds"])
            self.assertIsNone(metadata["finishedInTime"])
            self.assertEqual(metadata["recordingTimingCandidateSeconds"], 100)
            self.assertEqual(metadata["productReportedDurationSeconds"], 7)
            self.assertEqual(metadata["timingStatus"], "pending-measurement")

    def test_metadata_integration_uses_the_real_writer_without_publishing_its_duration(self):
        engine = load_engine()
        runner = PublicRecordingRunner(engine)
        runner.install()
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            recording = video(root)
            output = root / "output.xlsx"
            output.write_bytes(b"original frozen output; no Excel is opened")
            runner.turns = observed_turns()
            task = engine.TaskSpec("sample", root, {}, root / "input.xlsx", ("one", "two"), ("a", "b"), 300)
            runner.write_metadata(task=task, task_output_dir=root, product="copilot", tool="test",
                status="success", model="test", reasoning="test", started_at=None, finished_at="test",
                duration_seconds=140, agent_metadata={}, turns=runner.turns,
                transcript_capture=None, transcript_written=True, evaluation=None,
                screen_recording=recording, error=None, workbook_left_open=False,
                product_reported_turn_durations=[3, 4])
            metadata = timing.read_json(root / "metadata.json")
            self.assertIsNone(metadata["durationMs"])
            self.assertEqual(metadata["productReportedDurationSeconds"], 7)
            self.assertEqual(metadata["runnerObservedDurationSeconds"], 140)
            self.assertEqual(metadata["recordingTimingCandidateSeconds"], 100)
            self.assertEqual(metadata["captureAutomation"]["schemaVersion"], "atomic-capture-automation/v1")
            self.assertEqual(metadata["captureAutomation"]["outputWorkbookSha256"],
                             hashlib.sha256(output.read_bytes()).hexdigest())

    def test_two_turn_public_lifecycle_records_candidates_with_no_live_product_calls(self):
        engine = load_engine()
        runner = PublicRecordingRunner(engine)
        runner.install()
        runner.original["log"] = Mock()
        submitted = iter([1_010_000, 1_090_000])
        def operation(command, *args, **kwargs):
            if command == "send" and not kwargs.get("--stage-only"):
                return {"status": "ok", "submittedAtEpochMs": next(submitted), "finalAnswerCountBefore": 0}
            return {"status": "ok", "text": "completed answer"}
        runner.original["osascript_op"] = operation
        runner.original["wait_for_turn"] = Mock(side_effect=[
            {"status": "ok", "lastObservedAtEpochMs": 1_050_000, "elapsedSeconds": 44},
            {"status": "ok", "lastObservedAtEpochMs": 1_150_000, "elapsedSeconds": 65}])
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            workbook = root / "input.xlsx"
            workbook.write_bytes(b"fixture; no Excel session")
            task = engine.TaskSpec("sample", root, {}, workbook, ("one", "two"), ("a", "b"), 300)
            args = engine.build_parser().parse_args(["--record", "--no-sheet-tour", "--record-postroll", "0",
                                                     "--model", "test", "--reasoning", "high"])
            args.capture = False
            saved_recording = {}
            def start_recording(**kwargs):
                saved_recording.update(video(kwargs["task_output_dir"]))
                return SimpleNamespace(started_mono=engine.time.monotonic()), None
            runner.original["start_task_recording"] = start_recording
            with patch.object(engine, "request_excel_workbook_open", return_value={"status": "ok"}), \
                 patch.object(engine, "wait_for_recording_ready", return_value={"status": "recording"}), \
                 patch.object(engine, "verify_recording_before_submission", return_value={"status": "ok"}), \
                 patch.object(engine, "stop_task_recording", side_effect=lambda *a, **k: saved_recording), \
                 patch.object(engine, "write_transcript", return_value=True), \
                 patch.object(engine, "run_deterministic_evaluation", return_value={"ok": False, "error": "test fixture"}):
                result = runner.run_task(task, root / "public-run", args)
            self.assertEqual(result["status"], "success")
            self.assertIsNone(result["durationSeconds"])
            self.assertEqual(result["recordingTimingCandidateSeconds"], 100)
            self.assertEqual(result["timingStatus"], "pending-measurement")
            metadata = timing.read_json(root / "public-run/tasks/sample/metadata.json")
            self.assertEqual(metadata["runnerObservedDurationSeconds"], 100)

    def test_resume_preserves_failed_completed_and_interrupted_attempts(self):
        engine = load_engine()
        runner = PublicRecordingRunner(engine)
        runner.original["run_task"] = Mock(side_effect=AssertionError("must not execute"))
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            task = SimpleNamespace(task_id="sample")
            directory = root / "tasks/sample"
            directory.mkdir(parents=True)
            (directory / "output.xlsx").write_bytes(b"existing attempt")
            with self.assertRaisesRegex(ValueError, "Interrupted"):
                runner.run_task(task, root, SimpleNamespace())
            for status in ["success", "timeout", "error"]:
                metadata = {"taskId": "sample", "status": status, "timingPolicy": timing.POLICY}
                timing.write_json(directory / "metadata.json", metadata)
                result = runner.run_task(task, root, SimpleNamespace())
                self.assertTrue(result["resumedWithoutExecution"])
                self.assertEqual(result["status"], status)
                self.assertEqual(timing.read_json(directory / "metadata.json"), metadata)

    def test_public_manifest_refuses_other_runs_and_other_active_public_batch(self):
        engine = load_engine()
        runner = PublicRecordingRunner(engine)
        args = SimpleNamespace(lane="atomic", record=True, resume=True)
        runner.original["write_manifest"] = Mock()
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            timing.write_json(root / "manifest.json", {"id": "other"})
            before = (root / "manifest.json").read_bytes()
            with self.assertRaisesRegex(ValueError, "another timing policy"):
                runner.write_manifest(root, args)
            self.assertEqual((root / "manifest.json").read_bytes(), before)
            timing.write_json(root / "manifest.json", {"publicRunnerSchemaVersion": timing.PUBLIC_SCHEMA,
                                                       "timingPolicy": timing.POLICY})
            (root / ".public-run.lock").write_text("another active process")
            with self.assertRaises(FileExistsError):
                runner.write_manifest(root, args)
            runner.release_lock()
            self.assertEqual((root / ".public-run.lock").read_text(), "another active process")

def time_deadline(engine, seconds):
    return engine.time.monotonic() + seconds

if __name__ == "__main__":
    unittest.main()
