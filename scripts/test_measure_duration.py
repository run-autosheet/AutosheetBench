"""Tests for measuring a task's duration on its recording."""
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_duration as duration  # noqa: E402

IDLE = bytes([240] * 100)
TYPED = bytes([240] * 90 + [120] * 10)            # a typed message changes the Send button
BUSY = bytes([240] * 60 + [30] * 40)              # Stop button
STARTING = bytes([0] * 100)                        # a bright look while the product starts
FAINT_BUSY = bytes([240] * 75 + [150] * 25)        # Stop button without the pointer over it
POINTER_ON_IDLE = bytes([240] * 95 + [140] * 5)    # the pointer resting on Send after the reply
DIALOG = bytes([200] * 60 + [255] * 40)            # a dialog dimming the pane after the reply


def timeline(*segments, step=0.2):
    """Frames and times from (look, seconds) segments."""
    frames, pts, t = [], [], 0.0
    for look, seconds in segments:
        for _ in range(round(seconds / step)):
            frames.append(look)
            pts.append(round(t, 6))
            t += step
    return frames, pts


class MeasureTest(unittest.TestCase):
    def test_one_message_ends_on_the_first_idle_frame(self):
        frames, pts = timeline((TYPED, 2), (BUSY, 10), (IDLE, 5))
        turns, total = duration.measure(frames, pts, [2.0], refine=False)
        self.assertEqual(total, 10.0)
        self.assertEqual((turns[0]["startSeconds"], turns[0]["endSeconds"]), (2.0, 12.0))

    def test_messages_are_summed_and_waits_between_them_do_not_count(self):
        frames, pts = timeline((TYPED, 2), (BUSY, 8), (IDLE, 6), (TYPED, 2), (BUSY, 4), (IDLE, 3))
        turns, total = duration.measure(frames, pts, [2.0, 18.0], refine=False)
        self.assertEqual([turn["seconds"] for turn in turns], [8.0, 4.0])
        self.assertEqual(total, 12.0)

    def test_a_pause_inside_the_reply_counts(self):
        frames, pts = timeline((TYPED, 2), (BUSY, 5), (IDLE, 3), (BUSY, 4), (IDLE, 4))
        self.assertEqual(duration.measure(frames, pts, [2.0], refine=False)[1], 12.0)

    def test_a_starting_look_does_not_hide_the_fainter_busy_look(self):
        frames, pts = timeline((TYPED, 2), (STARTING, 2), (FAINT_BUSY, 10), (IDLE, 4))
        self.assertEqual(duration.measure(frames, pts, [2.0], refine=False)[1], 12.0)

    def test_the_pointer_resting_on_send_is_idle(self):
        frames, pts = timeline((TYPED, 2), (BUSY, 6), (POINTER_ON_IDLE, 20), (IDLE, 4))
        self.assertEqual(duration.measure(frames, pts, [2.0], refine=False)[1], 6.0)

    def test_a_dialog_over_the_pane_after_the_reply_is_not_work(self):
        frames, pts = timeline((TYPED, 2), (BUSY, 6), (IDLE, 10), (DIALOG, 5), (IDLE, 2))
        self.assertEqual(duration.measure(frames, pts, [2.0], refine=False)[1], 6.0)

    def test_the_next_message_being_typed_is_not_part_of_the_reply(self):
        frames, pts = timeline((TYPED, 2), (BUSY, 6), (IDLE, 10), (STARTING, 0.4), (TYPED, 1.6), (BUSY, 4), (IDLE, 3))
        self.assertEqual([turn["seconds"] for turn in duration.measure(frames, pts, [2.0, 20.0], refine=False)[0]], [6.0, 4.0])

    def test_reviewed_moments_settle_an_ambiguous_look(self):
        frames, pts = timeline((TYPED, 2), (STARTING, 8), (FAINT_BUSY, 10), (IDLE, 4))
        self.assertEqual(duration.measure(frames, pts, [2.0], refine=False)[1], 8.0)
        self.assertEqual(duration.measure(frames, pts, [2.0], refine=False, busy_at=15)[1], 18.0)
        frames, pts = timeline((TYPED, 2), (BUSY, 2), (POINTER_ON_IDLE, 2), (BUSY, 2), (POINTER_ON_IDLE, 20))
        self.assertEqual(duration.measure(frames, pts, [2.0], refine=False, ends_idle=False)[1], None)
        self.assertEqual(duration.measure(frames, pts, [2.0], refine=False, idle_at=[10])[1], 6.0)

    def test_an_approximate_send_time_is_refined_on_the_frames(self):
        frames, pts = timeline((TYPED, 5), (BUSY, 10), (IDLE, 5))
        for given in (3.4, 5.0, 6.6):
            turns, total = duration.measure(frames, pts, [given], refine=True)
            self.assertAlmostEqual(turns[0]["startSeconds"], 4.9, places=6)  # between the last typed and first busy frame
            self.assertAlmostEqual(total, 10.1, places=6)

    def test_a_send_time_given_too_late_is_reported(self):
        frames, pts = timeline((TYPED, 2), (BUSY, 10), (IDLE, 5))
        self.assertAlmostEqual(duration.measure(frames, pts, [4.0], refine=True)[1], 10.1)  # within the window
        turns, total = duration.measure(frames, pts, [6.0], refine=True)  # 4 s late: beyond it
        self.assertIsNone(total)
        self.assertIn("send time", turns[0]["problem"])

    def test_still_busy_at_the_end_has_no_end(self):
        frames, pts = timeline((TYPED, 2), (BUSY, 10))
        turns, total = duration.measure(frames, pts, [2.0], refine=False, ends_idle=False)
        self.assertIsNone(total)
        self.assertIn("still busy", turns[0]["problem"])

    def test_without_a_send_time_the_send_is_found_on_the_frames(self):
        frames, pts = timeline((IDLE, 3), (BUSY, 10), (IDLE, 5))
        sends = duration.find_sends(frames, pts, 1)
        self.assertEqual(sends, [3.0])
        self.assertAlmostEqual(duration.measure(frames, pts, sends)[1], 10.1)  # sent between 2.8 s and 3.0 s
        with self.assertRaises(ValueError):
            duration.find_sends(*timeline((IDLE, 10)), 1)

    def test_messages_sent_counts_fixed_replies_from_the_transcript(self):
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            self.assertIsNone(duration.messages_sent(folder))
            (folder / "message-01.txt").write_text("one")
            self.assertEqual(duration.messages_sent(folder), 1)
            (folder / "transcript.json").write_text(json.dumps({"entries": [
                {"role": "user", "text": "one"}, {"role": "assistant", "text": "Which sheet?"},
                {"role": "user", "text": "fixed reply"}, {"role": "assistant", "text": "Done"}]}))
            self.assertEqual(duration.messages_sent(folder), 2)

    def test_a_capture_gap_at_a_boundary_makes_the_turn_approximate(self):
        frames, pts = timeline((TYPED, 2), (BUSY, 6))
        frames.append(IDLE)
        pts.append(pts[-1] + 3.0)
        frames, pts = frames + [IDLE] * 5, pts + [pts[-1] + 0.2 * i for i in range(1, 6)]
        turns, _ = duration.measure(frames, pts, [2.0], refine=False)
        self.assertEqual(turns[0]["approximateBoundSeconds"], 3.0)


@unittest.skipUnless(shutil.which("ffmpeg") and shutil.which("ffprobe"), "needs ffmpeg")
class RecordingTest(unittest.TestCase):
    def test_measures_a_recording_and_writes_metadata(self):
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            video = folder / "recording.mp4"
            # A 64x64 video: a dark square (busy) from 2 s to 9 s, white otherwise.
            subprocess.run(["ffmpeg", "-v", "error", "-f", "lavfi", "-i", "color=white:s=64x64:r=5:d=12",
                            "-vf", "drawbox=x=40:y=40:w=16:h=16:color=black:t=fill:enable='between(t,2,9)'",
                            "-pix_fmt", "yuv420p", str(video)], check=True)
            (folder / "metadata.json").write_text(json.dumps({"timeLimitSeconds": 300}))
            args = ["--recording", str(video), "--region", "36,36,24,24", "--sent", "2.5", "--write"]
            self.assertEqual(duration.main(args), 0)
            report = json.loads((folder / "duration.json").read_text())
            self.assertAlmostEqual(report["seconds"], 7.3, delta=0.05)  # 1.9 s (between frames) to 9.2 s
            self.assertTrue(all(Path(file).exists() for file in report["reviewFrames"]))
            metadata = json.loads((folder / "metadata.json").read_text())
            self.assertEqual(metadata["timingStatus"], "measured")
            timing = metadata["timingCheck"]
            self.assertEqual(timing["recordingPromptToCompletionSeconds"], report["seconds"])
            self.assertEqual(timing["recordingTurns"], 1)
            with self.assertRaises(SystemExit):
                duration.main(args)  # a duration is written once unless --replace
            # One message sent: the send is found on the recording.
            (folder / "message-01.txt").write_text("the task")
            found = ["--recording", str(video), "--region", "36,36,24,24"]
            self.assertEqual(duration.main(found), 0)
            self.assertAlmostEqual(json.loads((folder / "duration.json").read_text())["seconds"], 7.3, delta=0.05)
            (folder / "message-02.txt").write_text("a follow-up")
            with self.assertRaises(SystemExit):
                duration.main(found)  # several messages need their send times


if __name__ == "__main__":
    unittest.main()
