// Records the Excel window: original video timeline plus host/wall-clock calibration.
#import <AVFoundation/AVFoundation.h>
#import <AppKit/AppKit.h>
#import <CoreMedia/CoreMedia.h>
#import <Foundation/Foundation.h>
#import <ScreenCaptureKit/ScreenCaptureKit.h>

static NSString *ArgumentValue(NSArray<NSString *> *arguments, NSString *key) {
    NSUInteger index = [arguments indexOfObject:key];
    if (index == NSNotFound || index + 1 >= arguments.count) return nil;
    return arguments[index + 1];
}

static NSString *NormalizedTitle(NSString *value) {
    return [[value ?: @"" lowercaseString] stringByReplacingOccurrencesOfString:@".xlsx" withString:@""];
}

static void WriteStatus(NSURL *url, NSDictionary *payload) {
    if (!url || ![NSJSONSerialization isValidJSONObject:payload]) return;
    NSData *data = [NSJSONSerialization dataWithJSONObject:payload
                                                   options:NSJSONWritingPrettyPrinted | NSJSONWritingSortedKeys
                                                     error:nil];
    [data writeToURL:url options:NSDataWritingAtomic error:nil];
}

static const size_t kSignatureGridWidth = 80;
static const size_t kSignatureGridHeight = 45;
static const NSInteger kSignatureChannelThreshold = 16;
static const double kSignatureChangedFractionThreshold = 0.006;
static const double kSignatureMeanDifferenceThreshold = 0.75;

static NSData *FrameSignature(CVPixelBufferRef pixelBuffer, uint64_t *hashOut) {
    if (!pixelBuffer ||
        CVPixelBufferLockBaseAddress(pixelBuffer, kCVPixelBufferLock_ReadOnly) != kCVReturnSuccess) {
        return nil;
    }
    const uint8_t *base = CVPixelBufferGetBaseAddress(pixelBuffer);
    size_t width = CVPixelBufferGetWidth(pixelBuffer);
    size_t height = CVPixelBufferGetHeight(pixelBuffer);
    size_t bytesPerRow = CVPixelBufferGetBytesPerRow(pixelBuffer);
    NSMutableData *signature = [NSMutableData dataWithLength:kSignatureGridWidth * kSignatureGridHeight * 2];
    uint8_t *destination = signature.mutableBytes;
    uint64_t hash = 1469598103934665603ULL;
    if (base && width > 0 && height > 0) {
        size_t regionTop = MIN(height - 1, (size_t)llround(height * 0.21));
        size_t regionBottom = MIN(height, MAX(regionTop + 1, (size_t)llround(height * 0.92)));
        size_t regionHeight = regionBottom - regionTop;
        for (size_t gy = 0; gy < kSignatureGridHeight; gy++) {
            size_t y = MIN(regionBottom - 1,
                           regionTop + ((2 * gy + 1) * regionHeight) / (2 * kSignatureGridHeight));
            const uint8_t *row = base + y * bytesPerRow;
            for (size_t gx = 0; gx < kSignatureGridWidth; gx++) {
                size_t x = MIN(width - 1, ((2 * gx + 1) * width) / (2 * kSignatureGridWidth));
                size_t rightX = MIN(width - 1, x + MAX((size_t)1, width / (2 * kSignatureGridWidth)));
                size_t lowerY = MIN(regionBottom - 1, y + MAX((size_t)1, regionHeight / (2 * kSignatureGridHeight)));
                const uint8_t *pixel = row + x * 4;
                const uint8_t *rightPixel = row + rightX * 4;
                const uint8_t *lowerPixel = base + lowerY * bytesPerRow + x * 4;
                NSInteger luminance = (29 * pixel[0] + 150 * pixel[1] + 77 * pixel[2]) >> 8;
                NSInteger rightLuminance =
                    (29 * rightPixel[0] + 150 * rightPixel[1] + 77 * rightPixel[2]) >> 8;
                NSInteger lowerLuminance =
                    (29 * lowerPixel[0] + 150 * lowerPixel[1] + 77 * lowerPixel[2]) >> 8;
                size_t offset = (gy * kSignatureGridWidth + gx) * 2;
                destination[offset] = (uint8_t)MIN(255, labs(luminance - rightLuminance));
                destination[offset + 1] = (uint8_t)MIN(255, labs(luminance - lowerLuminance));
                for (size_t component = 0; component < 2; component++) {
                    hash ^= destination[offset + component];
                    hash *= 1099511628211ULL;
                }
            }
        }
    }
    CVPixelBufferUnlockBaseAddress(pixelBuffer, kCVPixelBufferLock_ReadOnly);
    if (hashOut) *hashOut = hash;
    return signature;
}

static NSDictionary *DifferenceMetrics(NSData *candidate, NSData *reference) {
    if (!candidate || candidate.length != reference.length || candidate.length == 0) {
        return @{ @"changedFraction": @1.0, @"meanAbsoluteDifference": @255.0 };
    }
    const uint8_t *left = candidate.bytes;
    const uint8_t *right = reference.bytes;
    NSUInteger sampleCount = candidate.length;
    NSUInteger changed = 0;
    uint64_t totalDifference = 0;
    for (NSUInteger sample = 0; sample < sampleCount; sample++) {
        NSInteger difference = labs((long)left[sample] - (long)right[sample]);
        totalDifference += (uint64_t)difference;
        if (difference >= kSignatureChannelThreshold) changed += 1;
    }
    return @{
        @"changedFraction": @((double)changed / (double)sampleCount),
        @"meanAbsoluteDifference": @((double)totalDifference / (double)sampleCount),
    };
}

static NSDictionary *VisualCompletionSummary(
    NSArray<NSData *> *signatures,
    NSArray<NSNumber *> *timestamps
) {
    if (signatures.count == 0 || signatures.count != timestamps.count) return @{};
    NSData *finalSignature = signatures.lastObject;
    double firstTimestamp = timestamps.firstObject.doubleValue;
    double lastTimestamp = timestamps.lastObject.doubleValue;
    NSInteger lastDifferentIndex = -1;
    NSDictionary *boundaryMetrics = nil;
    double maximumTailChangedFraction = 0;
    double maximumTailMeanDifference = 0;
    for (NSInteger index = (NSInteger)signatures.count - 1; index >= 0; index--) {
        NSDictionary *metrics = DifferenceMetrics(signatures[(NSUInteger)index], finalSignature);
        double changedFraction = [metrics[@"changedFraction"] doubleValue];
        double meanDifference = [metrics[@"meanAbsoluteDifference"] doubleValue];
        BOOL materiallyDifferent =
            (changedFraction >= kSignatureChangedFractionThreshold && meanDifference >= 0.65) ||
            meanDifference >= kSignatureMeanDifferenceThreshold;
        if (materiallyDifferent) {
            lastDifferentIndex = index;
            boundaryMetrics = metrics;
            break;
        }
        maximumTailChangedFraction = MAX(maximumTailChangedFraction, changedFraction);
        maximumTailMeanDifference = MAX(maximumTailMeanDifference, meanDifference);
    }
    NSUInteger completionIndex = MIN((NSUInteger)(lastDifferentIndex + 1), signatures.count - 1);
    double completionSeconds = timestamps[completionIndex].doubleValue - firstTimestamp;
    double decodedDuration = MAX(0, lastTimestamp - firstTimestamp);
    return @{
        @"visualCompletionSeconds": @(completionSeconds),
        @"finalStableTailSeconds": @(MAX(0, decodedDuration - completionSeconds)),
        @"lastDifferentFrameSeconds": lastDifferentIndex >= 0
            ? @(timestamps[(NSUInteger)lastDifferentIndex].doubleValue - firstTimestamp)
            : [NSNull null],
        @"maximumTailChangedFraction": @(maximumTailChangedFraction),
        @"maximumTailMeanAbsoluteDifference": @(maximumTailMeanDifference),
        @"boundaryChangedFraction": boundaryMetrics ? boundaryMetrics[@"changedFraction"] : [NSNull null],
        @"boundaryMeanAbsoluteDifference": boundaryMetrics ? boundaryMetrics[@"meanAbsoluteDifference"] : [NSNull null],
        @"comparison": @"backward-from-final-frame",
        @"comparisonRegion": @"excel-content-below-ribbon",
        @"gridWidth": @(kSignatureGridWidth),
        @"gridHeight": @(kSignatureGridHeight),
        @"channelThreshold": @(kSignatureChannelThreshold),
        @"changedFractionThreshold": @(kSignatureChangedFractionThreshold),
        @"meanDifferenceThreshold": @(kSignatureMeanDifferenceThreshold),
        @"minimumMeanDifferenceForChangedFraction": @0.65,
    };
}

@interface RecordingDelegate : NSObject <SCRecordingOutputDelegate, SCStreamDelegate, SCStreamOutput>
@property(nonatomic) BOOL started;
@property(nonatomic) BOOL finished;
@property(nonatomic, copy) NSString *failure;
@property(nonatomic, copy) NSString *firstFramePath;
@property(nonatomic, copy) NSString *snapshotRequestPath;
@property(nonatomic) BOOL savedFirstFrame;
@property(nonatomic) NSUInteger frameCount;
@property(nonatomic) NSUInteger distinctFrameCount;
@property(nonatomic) uint64_t lastFrameHash;
@property(nonatomic) BOOL hasFrameHash;
@property(nonatomic) BOOL hasFirstFrameTimestamp;
@property(nonatomic) double firstFrameTimestamp;
@property(nonatomic) double firstFrameEpochMs;
@property(nonatomic) double clockMappingUncertaintyMs;
@property(nonatomic, strong) NSMutableArray<NSData *> *frameSignatures;
@property(nonatomic, strong) NSMutableArray<NSNumber *> *frameTimestamps;
@property(nonatomic, strong) NSMutableArray<NSNumber *> *frameChangeTimestamps;
@end

@implementation RecordingDelegate
- (void)recordingOutputDidStartRecording:(SCRecordingOutput *)recordingOutput {
    @synchronized(self) { self.started = YES; }
}
- (void)recordingOutput:(SCRecordingOutput *)recordingOutput didFailWithError:(NSError *)error {
    @synchronized(self) {
        self.failure = error.localizedDescription;
        self.finished = YES;
    }
}
- (void)recordingOutputDidFinishRecording:(SCRecordingOutput *)recordingOutput {
    @synchronized(self) { self.finished = YES; }
}
- (void)stream:(SCStream *)stream didStopWithError:(NSError *)error {
    @synchronized(self) {
        self.failure = error.localizedDescription;
        self.finished = YES;
    }
}
- (void)stream:(SCStream *)stream
    didOutputSampleBuffer:(CMSampleBufferRef)sampleBuffer
                   ofType:(SCStreamOutputType)type {
    if (type != SCStreamOutputTypeScreen || !CMSampleBufferIsValid(sampleBuffer)) return;
    CVPixelBufferRef pixelBuffer = CMSampleBufferGetImageBuffer(sampleBuffer);
    uint64_t hash = 0;
    NSData *signature = FrameSignature(pixelBuffer, &hash);
    double timestamp = CMTimeGetSeconds(CMSampleBufferGetPresentationTimeStamp(sampleBuffer));
    if (!signature || !isfinite(timestamp)) return;
    @synchronized(self) {
        // Capture the same isolated window as the video, without activating it.
        BOOL snapshotRequested = self.snapshotRequestPath.length &&
            [[NSFileManager defaultManager] fileExistsAtPath:self.snapshotRequestPath];
        if (self.firstFramePath.length && (!self.savedFirstFrame || snapshotRequested) &&
            CVPixelBufferLockBaseAddress(pixelBuffer, kCVPixelBufferLock_ReadOnly) == kCVReturnSuccess) {
            CGColorSpaceRef colorSpace = CGColorSpaceCreateDeviceRGB();
            CGContextRef context = CGBitmapContextCreate(CVPixelBufferGetBaseAddress(pixelBuffer),
                CVPixelBufferGetWidth(pixelBuffer), CVPixelBufferGetHeight(pixelBuffer), 8,
                CVPixelBufferGetBytesPerRow(pixelBuffer), colorSpace,
                kCGBitmapByteOrder32Little | kCGImageAlphaPremultipliedFirst);
            if (context) {
                CGImageRef image = CGBitmapContextCreateImage(context);
                if (image) {
                    NSBitmapImageRep *bitmap = [[NSBitmapImageRep alloc] initWithCGImage:image];
                    NSData *png = [bitmap representationUsingType:NSBitmapImageFileTypePNG properties:@{}];
                    BOOL saved = [png writeToFile:self.firstFramePath atomically:YES];
                    self.savedFirstFrame = self.savedFirstFrame || saved;
                    if (saved && snapshotRequested) {
                        [[NSFileManager defaultManager] removeItemAtPath:self.snapshotRequestPath error:nil];
                    }
                    CGImageRelease(image);
                }
                CGContextRelease(context);
            }
            CGColorSpaceRelease(colorSpace);
            CVPixelBufferUnlockBaseAddress(pixelBuffer, kCVPixelBufferLock_ReadOnly);
        }
        if (!self.frameSignatures) self.frameSignatures = [NSMutableArray array];
        if (!self.frameTimestamps) self.frameTimestamps = [NSMutableArray array];
        if (!self.frameChangeTimestamps) self.frameChangeTimestamps = [NSMutableArray array];
        if (!self.hasFirstFrameTimestamp) {
            self.firstFrameTimestamp = timestamp;
            double wallBeforeMs = [NSDate date].timeIntervalSince1970 * 1000.0;
            double hostNow = CMTimeGetSeconds(CMClockGetTime(CMClockGetHostTimeClock()));
            double wallAfterMs = [NSDate date].timeIntervalSince1970 * 1000.0;
            self.firstFrameEpochMs = (wallBeforeMs + wallAfterMs) / 2.0
                + (timestamp - hostNow) * 1000.0;
            self.clockMappingUncertaintyMs = wallAfterMs - wallBeforeMs;
            self.hasFirstFrameTimestamp = YES;
        }
        double relativeTimestamp = MAX(0, timestamp - self.firstFrameTimestamp);
        self.frameCount += 1;
        [self.frameSignatures addObject:signature];
        [self.frameTimestamps addObject:@(timestamp)];
        if (!self.hasFrameHash || hash != self.lastFrameHash) {
            self.distinctFrameCount += 1;
            self.lastFrameHash = hash;
            self.hasFrameHash = YES;
            [self.frameChangeTimestamps addObject:@(relativeTimestamp)];
        }
    }
}
@end

static SCShareableContent *ShareableContent(NSError **error) {
    __block SCShareableContent *result = nil;
    __block NSError *capturedError = nil;
    dispatch_semaphore_t semaphore = dispatch_semaphore_create(0);
    [SCShareableContent getShareableContentExcludingDesktopWindows:NO
                                              onScreenWindowsOnly:NO
                                               completionHandler:^(SCShareableContent *content, NSError *contentError) {
        result = content;
        capturedError = contentError;
        dispatch_semaphore_signal(semaphore);
    }];
    dispatch_semaphore_wait(semaphore, dispatch_time(DISPATCH_TIME_NOW, 15 * NSEC_PER_SEC));
    if (error) *error = capturedError;
    return result;
}

static SCWindow *FindExcelWindow(
    NSString *title,
    NSString *applicationBundle,
    CGWindowID requestedWindowId,
    NSTimeInterval timeout,
    SCShareableContent **matchedContent,
    NSError **error
) {
    NSString *wanted = NormalizedTitle(title);
    NSDate *deadline = [NSDate dateWithTimeIntervalSinceNow:timeout];
    do {
        NSError *contentError = nil;
        SCShareableContent *content = ShareableContent(&contentError);
        if (contentError) {
            if (error) *error = contentError;
            return nil;
        }
        for (SCWindow *window in content.windows) {
            NSString *bundle = window.owningApplication.bundleIdentifier.lowercaseString ?: @"";
            NSString *owner = window.owningApplication.applicationName.lowercaseString ?: @"";
            NSString *candidate = NormalizedTitle(window.title);
            BOOL isExcel = [bundle isEqualToString:applicationBundle.lowercaseString] ||
                           ([applicationBundle isEqualToString:@"com.microsoft.Excel"] &&
                            [owner isEqualToString:@"microsoft excel"]);
            BOOL usableSize = window.frame.size.width >= 200 && window.frame.size.height >= 120;
            if (isExcel && usableSize &&
                (requestedWindowId ? window.windowID == requestedWindowId :
                 ([candidate isEqualToString:wanted] || [candidate containsString:wanted]))) {
                if (matchedContent) *matchedContent = content;
                return window;
            }
        }
        [NSThread sleepForTimeInterval:0.25];
    } while ([deadline timeIntervalSinceNow] > 0);
    if (error) {
        *error = [NSError errorWithDomain:@"ExcelWindowRecorder"
                                     code:2
                                 userInfo:@{NSLocalizedDescriptionKey:
                                                [NSString stringWithFormat:@"Window not found for %@: %@", applicationBundle, title]}];
    }
    return nil;
}

static BOOL WaitForFlag(RecordingDelegate *delegate, NSString *property, NSTimeInterval timeout) {
    NSDate *deadline = [NSDate dateWithTimeIntervalSinceNow:timeout];
    do {
        @synchronized(delegate) {
            if (delegate.failure.length > 0) return NO;
            if ([property isEqualToString:@"started"] && delegate.started) return YES;
            if ([property isEqualToString:@"finished"] && delegate.finished) return YES;
            if ([property isEqualToString:@"hasFirstFrameTimestamp"] && delegate.hasFirstFrameTimestamp) return YES;
        }
        [[NSRunLoop currentRunLoop] runUntilDate:[NSDate dateWithTimeIntervalSinceNow:0.05]];
    } while ([deadline timeIntervalSinceNow] > 0);
    return NO;
}

int main(int argc, const char *argv[]) {
    @autoreleasepool {
        [NSApplication sharedApplication];
        [NSApp setActivationPolicy:NSApplicationActivationPolicyProhibited];
        NSArray<NSString *> *arguments = NSProcessInfo.processInfo.arguments;
        NSString *windowTitle = ArgumentValue(arguments, @"--window-title");
        CGWindowID requestedWindowId = (CGWindowID)[ArgumentValue(arguments, @"--window-id") longLongValue];
        NSString *applicationBundle = ArgumentValue(arguments, @"--app-bundle") ?: @"com.microsoft.Excel";
        BOOL fullDisplay = [arguments containsObject:@"--full-display"];
        BOOL isolatedWindow = [arguments containsObject:@"--isolated-window"];
        NSString *outputPath = ArgumentValue(arguments, @"--output");
        NSString *statusPath = ArgumentValue(arguments, @"--status-file");
        NSString *stopPath = ArgumentValue(arguments, @"--stop-file");
        NSTimeInterval windowTimeout = [ArgumentValue(arguments, @"--window-timeout") doubleValue] ?: 120;
        NSTimeInterval maxDuration = [ArgumentValue(arguments, @"--max-duration") doubleValue] ?: 900;
        NSInteger framesPerSecond = [ArgumentValue(arguments, @"--fps") integerValue] ?: 5;
        NSInteger maxWidth = [ArgumentValue(arguments, @"--max-width") integerValue] ?: 1920;
        NSURL *outputURL = outputPath ? [NSURL fileURLWithPath:outputPath] : nil;
        NSURL *statusURL = statusPath ? [NSURL fileURLWithPath:statusPath] : nil;
        NSDate *processStarted = [NSDate date];

        if (!windowTitle.length || !outputPath.length || !statusPath.length || !stopPath.length ||
            windowTimeout <= 0 || maxDuration <= 0 || framesPerSecond <= 0 || maxWidth <= 0 || (fullDisplay && isolatedWindow)) {
            fprintf(stderr, "Usage: record_window --window-title TITLE --output PATH --status-file PATH --stop-file PATH [--app-bundle ID] [--full-display | --isolated-window] [--first-frame PNG] [--snapshot-request PATH] [--window-timeout SEC] [--max-duration SEC] [--fps N] [--max-width PX]\n");
            return 2;
        }

        NSFileManager *fileManager = NSFileManager.defaultManager;
        [fileManager createDirectoryAtURL:outputURL.URLByDeletingLastPathComponent
              withIntermediateDirectories:YES attributes:nil error:nil];
        [fileManager removeItemAtURL:outputURL error:nil];
        WriteStatus(statusURL, @{ @"status": @"waiting-for-window", @"windowTitle": windowTitle });

        __block NSError *error = nil;
        SCShareableContent *content = nil;
        SCWindow *window = FindExcelWindow(windowTitle, applicationBundle, requestedWindowId, windowTimeout, &content, &error);
        if (!window) {
            WriteStatus(statusURL, @{ @"status": @"error",
                                      @"error": error.localizedDescription ?: @"Window not found" });
            fprintf(stderr, "record_window: %s\n", error.localizedDescription.UTF8String);
            return 1;
        }

        SCDisplay *display = nil;
        CGPoint windowCenter = CGPointMake(CGRectGetMidX(window.frame), CGRectGetMidY(window.frame));
        for (SCDisplay *candidate in content.displays) {
            if (CGRectContainsPoint(candidate.frame, windowCenter)) {
                display = candidate;
                break;
            }
        }
        if (!display) display = content.displays.firstObject;
        if (!content || !display || !window.owningApplication) {
            NSString *message = error.localizedDescription ?: @"Could not resolve the Excel window's display";
            WriteStatus(statusURL, @{ @"status": @"error", @"error": message });
            return 1;
        }
        // Excel add-in panes are hosted in a remotely rendered WebView. Both
        // window-only and Excel-application-only filters freeze that surface
        // at its initial frame. Capture the live display composition cropped
        // to this workbook window so WebView updates continue to arrive.
        // Chrome split views can be captured independently of foreground apps.
        // Retain display composition for Excel's remotely rendered add-in pane.
        SCContentFilter *filter = isolatedWindow
            ? [[SCContentFilter alloc] initWithDesktopIndependentWindow:window]
            : [[SCContentFilter alloc] initWithDisplay:display excludingWindows:@[]];
        SCStreamConfiguration *configuration = [[SCStreamConfiguration alloc] init];
        // A full display captures Claude Desktop and the live Sheets browser together.
        CGRect cropRect = isolatedWindow ? window.frame : (fullDisplay ? display.frame : CGRectIntersection(window.frame, display.frame));
        CGRect localCropRect = isolatedWindow ? CGRectMake(0, 0, window.frame.size.width, window.frame.size.height)
            : CGRectOffset(cropRect, -display.frame.origin.x, -display.frame.origin.y);
        // SCDisplay's width and frame are both in points, so their ratio is
        // always 1 and Retina windows were captured at half resolution. The
        // filter reports the display's real pixels per point (2 on Retina).
        CGFloat displayScale = 0;
        if (@available(macOS 14.0, *)) {
            displayScale = filter.pointPixelScale;
        }
        if (displayScale <= 0) {
            for (NSScreen *screen in NSScreen.screens) {
                if ([screen.deviceDescription[@"NSScreenNumber"] unsignedIntValue] == display.displayID) {
                    displayScale = screen.backingScaleFactor;
                }
            }
        }
        if (displayScale <= 0) {
            displayScale = 2.0;
        }
        NSInteger sourceWidth = MAX(2, (NSInteger)llround(cropRect.size.width * displayScale));
        NSInteger sourceHeight = MAX(2, (NSInteger)llround(cropRect.size.height * displayScale));
        double scale = MIN(1.0, (double)maxWidth / (double)sourceWidth);
        configuration.width = MAX(2, ((NSInteger)llround(sourceWidth * scale) / 2) * 2);
        configuration.height = MAX(2, ((NSInteger)llround(sourceHeight * scale) / 2) * 2);
        configuration.sourceRect = localCropRect;
        configuration.minimumFrameInterval = CMTimeMake(1, (int32_t)framesPerSecond);
        configuration.queueDepth = 3;
        configuration.showsCursor = YES;
        configuration.capturesAudio = NO;
        configuration.includeChildWindows = YES;
        configuration.ignoreGlobalClipDisplay = YES;
        configuration.pixelFormat = kCVPixelFormatType_32BGRA;

        RecordingDelegate *delegate = [[RecordingDelegate alloc] init];
        delegate.firstFramePath = ArgumentValue(arguments, @"--first-frame");
        delegate.snapshotRequestPath = ArgumentValue(arguments, @"--snapshot-request");
        SCStream *stream = [[SCStream alloc] initWithFilter:filter configuration:configuration delegate:delegate];
        dispatch_queue_t sampleQueue = dispatch_queue_create("benchmark.excel-recording.frames", DISPATCH_QUEUE_SERIAL);
        if (![stream addStreamOutput:delegate
                                type:SCStreamOutputTypeScreen
                  sampleHandlerQueue:sampleQueue
                               error:&error]) {
            WriteStatus(statusURL, @{ @"status": @"error", @"error": error.localizedDescription ?: @"Could not inspect recording frames" });
            return 1;
        }
        SCRecordingOutputConfiguration *outputConfiguration = [[SCRecordingOutputConfiguration alloc] init];
        outputConfiguration.outputURL = outputURL;
        outputConfiguration.videoCodecType = AVVideoCodecTypeH264;
        outputConfiguration.outputFileType = AVFileTypeMPEG4;
        SCRecordingOutput *recordingOutput = [[SCRecordingOutput alloc]
                                               initWithConfiguration:outputConfiguration
                                               delegate:delegate];
        if (![stream addRecordingOutput:recordingOutput error:&error]) {
            WriteStatus(statusURL, @{ @"status": @"error", @"error": error.localizedDescription ?: @"Could not add recording output" });
            return 1;
        }

        dispatch_semaphore_t startSemaphore = dispatch_semaphore_create(0);
        [stream startCaptureWithCompletionHandler:^(NSError *startError) {
            error = startError;
            dispatch_semaphore_signal(startSemaphore);
        }];
        dispatch_semaphore_wait(startSemaphore, dispatch_time(DISPATCH_TIME_NOW, 15 * NSEC_PER_SEC));
        if (error || !WaitForFlag(delegate, @"started", 10)) {
            NSString *message = error.localizedDescription ?: delegate.failure ?: @"ScreenCaptureKit did not start recording";
            WriteStatus(statusURL, @{ @"status": @"error", @"error": message });
            fprintf(stderr, "record_window: %s\n", message.UTF8String);
            return 1;
        }

        if (!WaitForFlag(delegate, @"hasFirstFrameTimestamp", 10)) {
            WriteStatus(statusURL, @{ @"status": @"error", @"error": @"No first video frame for timing calibration" });
            return 1;
        }
        WriteStatus(statusURL, @{
            @"timingSchemaVersion": @"atomic-recording-clock/v1",
            @"firstFrameEpochMs": @(delegate.firstFrameEpochMs),
            @"clockMappingUncertaintyMs": @(delegate.clockMappingUncertaintyMs),
            @"status": @"recording",
            @"captureMode": isolatedWindow ? @"window" : (fullDisplay ? @"display" : @"window-crop"),
            @"applicationBundle": applicationBundle,
            @"displayId": @(display.displayID),
            @"windowTitle": window.title ?: windowTitle,
            @"windowId": @(window.windowID),
            @"width": @(configuration.width),
            @"height": @(configuration.height),
            @"framesPerSecond": @(framesPerSecond),
        });

        NSDate *deadline = [NSDate dateWithTimeIntervalSinceNow:maxDuration];
        while (![fileManager fileExistsAtPath:stopPath] && [deadline timeIntervalSinceNow] > 0) {
            @synchronized(delegate) {
                if (delegate.failure.length > 0) break;
            }
            [NSThread sleepForTimeInterval:0.2];
        }

        __block BOOL stopCompleted = NO;
        [stream stopCaptureWithCompletionHandler:^(NSError *stopError) {
            if (stopError && !error) error = stopError;
            stopCompleted = YES;
        }];
        NSDate *stopDeadline = [NSDate dateWithTimeIntervalSinceNow:5];
        while (!stopCompleted && [stopDeadline timeIntervalSinceNow] > 0) {
            [[NSRunLoop currentRunLoop] runUntilDate:[NSDate dateWithTimeIntervalSinceNow:0.05]];
        }
        if (!stopCompleted && !error) {
            error = [NSError errorWithDomain:@"ExcelWindowRecorder"
                                         code:3
                                     userInfo:@{NSLocalizedDescriptionKey:
                                                    @"ScreenCaptureKit did not confirm recorder shutdown"}];
        }
        WaitForFlag(delegate, @"finished", 5);

        NSString *failure = error.localizedDescription ?: delegate.failure;
        NSUInteger frameCount = 0;
        NSUInteger distinctFrameCount = 0;
        NSArray<NSData *> *frameSignatures = nil;
        NSArray<NSNumber *> *frameTimestamps = nil;
        NSArray<NSNumber *> *frameChangeTimestamps = nil;
        @synchronized(delegate) {
            frameCount = delegate.frameCount;
            distinctFrameCount = delegate.distinctFrameCount;
            frameSignatures = [delegate.frameSignatures copy] ?: @[];
            frameTimestamps = [delegate.frameTimestamps copy] ?: @[];
            frameChangeTimestamps = [delegate.frameChangeTimestamps copy] ?: @[];
        }
        if (!failure.length && (frameCount < 2 || distinctFrameCount < 2)) {
            failure = [NSString stringWithFormat:
                @"Frozen recording detected (%lu frames, %lu distinct)",
                (unsigned long)frameCount,
                (unsigned long)distinctFrameCount];
        }
        if (failure.length > 0) {
            WriteStatus(statusURL, @{ @"status": @"error", @"error": failure,
                                      @"frameCount": @(frameCount),
                                      @"distinctFrameCount": @(distinctFrameCount) });
            fprintf(stderr, "record_window: %s\n", failure.UTF8String);
            return 1;
        }
        NSDictionary *attributes = [fileManager attributesOfItemAtPath:outputPath error:&error];
        NSNumber *bytes = attributes[NSFileSize] ?: @0;
        // SCRecordingOutput.recordedDuration can be rounded/truncated even
        // after shutdown. Read the finalized MP4's precise media duration;
        // neither that coarse counter nor wall time defines the video axis.
        AVURLAsset *recordedAsset = [AVURLAsset URLAssetWithURL:[NSURL fileURLWithPath:outputPath]
                                                      options:@{AVURLAssetPreferPreciseDurationAndTimingKey: @YES}];
        double recordedSeconds = CMTimeGetSeconds(recordedAsset.duration);
        double firstPts = frameTimestamps.count ? frameTimestamps.firstObject.doubleValue : 0;
        double lastOffset = frameTimestamps.count ? frameTimestamps.lastObject.doubleValue - firstPts : 0;
        if (!isfinite(recordedSeconds) || recordedSeconds <= 0 || recordedSeconds < lastOffset) {
            WriteStatus(statusURL, @{ @"status": @"error", @"error": @"Finalized MP4 duration does not cover captured frame timestamps" });
            return 1;
        }
        NSMutableDictionary *success = [@{
            @"status": @"success",
            @"captureMode": isolatedWindow ? @"window" : (fullDisplay ? @"display" : @"window-crop"),
            @"applicationBundle": applicationBundle,
            @"displayId": @(display.displayID),
            @"windowTitle": window.title ?: windowTitle,
            @"windowId": @(window.windowID),
            @"durationSeconds": @(recordedSeconds),
            @"durationSource": @"finalized-mp4-media-timeline",
            @"captureReportedDurationSeconds": @(CMTimeGetSeconds(recordingOutput.recordedDuration)),
            @"bytes": bytes,
            @"width": @(configuration.width),
            @"height": @(configuration.height),
            @"framesPerSecond": @(framesPerSecond),
            @"frameCount": @(frameCount),
            @"distinctFrameCount": @(distinctFrameCount),
            @"frameChangeTimestampsSeconds": frameChangeTimestamps,
        } mutableCopy];
        NSMutableArray<NSNumber *> *frameOffsets = [NSMutableArray array];
        for (NSNumber *pts in frameTimestamps) [frameOffsets addObject:@(pts.doubleValue - firstPts)];
        success[@"timingSchemaVersion"] = @"atomic-recording-clock/v1";
        success[@"firstFrameEpochMs"] = @(delegate.firstFrameEpochMs);
        success[@"clockMappingUncertaintyMs"] = @(delegate.clockMappingUncertaintyMs);
        success[@"firstFramePresentationTimeSeconds"] = @(firstPts);
        success[@"frameOffsetsSeconds"] = frameOffsets;
        success[@"timeCompressionApplied"] = @NO;
        [success addEntriesFromDictionary:VisualCompletionSummary(frameSignatures, frameTimestamps)];
        if (fullDisplay) success[@"comparisonRegion"] = @"display-content-below-header";
        WriteStatus(statusURL, success);
        return 0;
    }
}
