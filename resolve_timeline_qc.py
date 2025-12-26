#!/usr/bin/env python
# DaVinci Resolve - Timeline Quality Control
# ===========================================
# Checks timeline for common issues and generates a report.
#
# Checks performed:
# - Video gaps (frames with no video content)
# - Flash frames (very short clips)
# - Audio overlaps on same track
# - Audio gaps
# - Disabled/muted clips
# - Offline media
# - Clips at end of source media

# ============== CONFIG ==============
FLASH_FRAME_THRESHOLD = 3      # Clips shorter than this are flagged
CHECK_AUDIO_GAPS = True        # Set False if audio gaps are intentional
MIN_AUDIO_GAP_FRAMES = 2       # Ignore gaps smaller than this
IGNORE_TRACK_NAMES = []        # Track names to skip (e.g., ["Music", "SFX"])
# ====================================


def get_resolve():
    try:
        import DaVinciResolveScript as dvr
        return dvr.scriptapp("Resolve")
    except ImportError:
        return bmd.scriptapp("Resolve")


def frames_to_tc(frames, fps):
    fps = int(round(fps))
    if frames < 0:
        frames = 0
    f = frames % fps
    s = (frames // fps) % 60
    m = (frames // (fps * 60)) % 60
    h = frames // (fps * 3600)
    return "{:02d}:{:02d}:{:02d}:{:02d}".format(h, m, s, f)


def get_track_items_sorted(timeline, track_type, track_index):
    """Get all items on a track, sorted by start position"""
    items = timeline.GetItemListInTrack(track_type, track_index)
    if not items:
        return []

    item_list = []
    for item in items:
        item_list.append({
            'item': item,
            'name': item.GetName(),
            'start': item.GetStart(),
            'end': item.GetEnd(),
            'duration': item.GetDuration(),
        })

    item_list.sort(key=lambda x: x['start'])
    return item_list


def check_video_gaps(timeline, fps, timeline_start, timeline_end):
    """Check for gaps in video coverage"""
    issues = []
    video_track_count = timeline.GetTrackCount("video")

    if video_track_count == 0:
        return issues

    # Collect all video clip ranges across all tracks
    all_video_ranges = []

    for track_idx in range(1, video_track_count + 1):
        items = get_track_items_sorted(timeline, "video", track_idx)
        for item in items:
            all_video_ranges.append((item['start'], item['end']))

    if not all_video_ranges:
        issues.append({
            'type': 'VIDEO_GAP',
            'severity': 'ERROR',
            'start': timeline_start,
            'end': timeline_end,
            'duration': timeline_end - timeline_start,
            'message': 'No video clips on timeline'
        })
        return issues

    # Sort by start position
    all_video_ranges.sort(key=lambda x: x[0])

    # Merge overlapping ranges
    merged = []
    for start, end in all_video_ranges:
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))

    # Check for gaps
    # Gap at beginning
    if merged[0][0] > timeline_start:
        gap_duration = merged[0][0] - timeline_start
        issues.append({
            'type': 'VIDEO_GAP',
            'severity': 'ERROR',
            'start': timeline_start,
            'end': merged[0][0],
            'duration': gap_duration,
            'message': 'Gap at timeline start ({} frames)'.format(gap_duration)
        })

    # Gaps between clips
    for i in range(len(merged) - 1):
        if merged[i][1] < merged[i+1][0]:
            gap_start = merged[i][1]
            gap_end = merged[i+1][0]
            gap_duration = gap_end - gap_start
            issues.append({
                'type': 'VIDEO_GAP',
                'severity': 'ERROR',
                'start': gap_start,
                'end': gap_end,
                'duration': gap_duration,
                'message': 'Video gap ({} frames)'.format(gap_duration)
            })

    # Gap at end
    if merged[-1][1] < timeline_end:
        gap_duration = timeline_end - merged[-1][1]
        issues.append({
            'type': 'VIDEO_GAP',
            'severity': 'WARNING',
            'start': merged[-1][1],
            'end': timeline_end,
            'duration': gap_duration,
            'message': 'Gap at timeline end ({} frames)'.format(gap_duration)
        })

    return issues


def check_flash_frames(timeline, fps):
    """Check for very short clips (flash frames)"""
    issues = []

    # Check video tracks
    video_track_count = timeline.GetTrackCount("video")
    for track_idx in range(1, video_track_count + 1):
        items = get_track_items_sorted(timeline, "video", track_idx)
        for item in items:
            if item['duration'] < FLASH_FRAME_THRESHOLD:
                issues.append({
                    'type': 'FLASH_FRAME',
                    'severity': 'WARNING',
                    'start': item['start'],
                    'end': item['end'],
                    'duration': item['duration'],
                    'track': 'V{}'.format(track_idx),
                    'clip': item['name'],
                    'message': 'Flash frame on V{}: "{}" ({} frames)'.format(
                        track_idx, item['name'], item['duration'])
                })

    # Check audio tracks
    audio_track_count = timeline.GetTrackCount("audio")
    for track_idx in range(1, audio_track_count + 1):
        items = get_track_items_sorted(timeline, "audio", track_idx)
        for item in items:
            if item['duration'] < FLASH_FRAME_THRESHOLD:
                issues.append({
                    'type': 'FLASH_FRAME',
                    'severity': 'WARNING',
                    'start': item['start'],
                    'end': item['end'],
                    'duration': item['duration'],
                    'track': 'A{}'.format(track_idx),
                    'clip': item['name'],
                    'message': 'Flash frame on A{}: "{}" ({} frames)'.format(
                        track_idx, item['name'], item['duration'])
                })

    return issues


def check_audio_overlaps(timeline, fps):
    """Check for overlapping audio clips on the same track"""
    issues = []
    audio_track_count = timeline.GetTrackCount("audio")

    for track_idx in range(1, audio_track_count + 1):
        items = get_track_items_sorted(timeline, "audio", track_idx)

        for i in range(len(items) - 1):
            current = items[i]
            next_item = items[i + 1]

            # Check if current clip extends into next clip
            if current['end'] > next_item['start']:
                overlap = current['end'] - next_item['start']
                issues.append({
                    'type': 'AUDIO_OVERLAP',
                    'severity': 'ERROR',
                    'start': next_item['start'],
                    'end': current['end'],
                    'duration': overlap,
                    'track': 'A{}'.format(track_idx),
                    'message': 'Audio overlap on A{}: "{}" and "{}" ({} frames)'.format(
                        track_idx, current['name'], next_item['name'], overlap)
                })

    return issues


def check_audio_gaps(timeline, fps):
    """Check for gaps in audio tracks"""
    if not CHECK_AUDIO_GAPS:
        return []

    issues = []
    audio_track_count = timeline.GetTrackCount("audio")

    for track_idx in range(1, audio_track_count + 1):
        track_name = timeline.GetTrackName("audio", track_idx)
        if track_name in IGNORE_TRACK_NAMES:
            continue

        items = get_track_items_sorted(timeline, "audio", track_idx)
        if len(items) < 2:
            continue

        for i in range(len(items) - 1):
            current = items[i]
            next_item = items[i + 1]

            gap = next_item['start'] - current['end']
            if gap >= MIN_AUDIO_GAP_FRAMES:
                issues.append({
                    'type': 'AUDIO_GAP',
                    'severity': 'INFO',
                    'start': current['end'],
                    'end': next_item['start'],
                    'duration': gap,
                    'track': 'A{}'.format(track_idx),
                    'message': 'Audio gap on A{}: {} frames between clips'.format(
                        track_idx, gap)
                })

    return issues


def check_disabled_clips(timeline, fps):
    """Check for disabled or muted clips"""
    issues = []

    # Check video tracks for disabled clips
    video_track_count = timeline.GetTrackCount("video")
    for track_idx in range(1, video_track_count + 1):
        items = timeline.GetItemListInTrack("video", track_idx)
        if not items:
            continue
        for item in items:
            # Check if clip is disabled (GetEnabled returns False)
            try:
                # Try different property names that might indicate disabled state
                props = item.GetProperty()
                if props and isinstance(props, dict):
                    if props.get('Disabled') or props.get('enabled') == False:
                        issues.append({
                            'type': 'DISABLED_CLIP',
                            'severity': 'INFO',
                            'start': item.GetStart(),
                            'end': item.GetEnd(),
                            'duration': item.GetDuration(),
                            'track': 'V{}'.format(track_idx),
                            'clip': item.GetName(),
                            'message': 'Disabled clip on V{}: "{}"'.format(
                                track_idx, item.GetName())
                        })
            except:
                pass

    # Check for muted audio tracks
    audio_track_count = timeline.GetTrackCount("audio")
    for track_idx in range(1, audio_track_count + 1):
        try:
            # Check if entire track is muted
            is_muted = timeline.GetIsTrackEnabled("audio", track_idx) == False
            if is_muted:
                issues.append({
                    'type': 'MUTED_TRACK',
                    'severity': 'WARNING',
                    'start': 0,
                    'end': 0,
                    'duration': 0,
                    'track': 'A{}'.format(track_idx),
                    'message': 'Audio track A{} is muted/disabled'.format(track_idx)
                })
        except:
            pass

    return issues


def check_offline_media(timeline, fps):
    """Check for offline/missing media"""
    issues = []

    video_track_count = timeline.GetTrackCount("video")
    for track_idx in range(1, video_track_count + 1):
        items = timeline.GetItemListInTrack("video", track_idx)
        if not items:
            continue
        for item in items:
            try:
                media_pool_item = item.GetMediaPoolItem()
                if media_pool_item:
                    clip_props = media_pool_item.GetClipProperty()
                    if clip_props:
                        # Check for offline status
                        if clip_props.get('Offline') or clip_props.get('File Path') == '':
                            issues.append({
                                'type': 'OFFLINE_MEDIA',
                                'severity': 'ERROR',
                                'start': item.GetStart(),
                                'end': item.GetEnd(),
                                'duration': item.GetDuration(),
                                'track': 'V{}'.format(track_idx),
                                'clip': item.GetName(),
                                'message': 'Offline media on V{}: "{}"'.format(
                                    track_idx, item.GetName())
                            })
            except:
                pass

    return issues


def check_source_end(timeline, fps):
    """Check if clips are trimmed to the very end of source media"""
    issues = []

    video_track_count = timeline.GetTrackCount("video")
    for track_idx in range(1, video_track_count + 1):
        items = timeline.GetItemListInTrack("video", track_idx)
        if not items:
            continue
        for item in items:
            try:
                media_pool_item = item.GetMediaPoolItem()
                if media_pool_item:
                    clip_props = media_pool_item.GetClipProperty()
                    if clip_props:
                        # Get source duration and check if we're at the end
                        source_frames = clip_props.get('Frames')
                        if source_frames:
                            source_frames = int(source_frames)
                            right_offset = item.GetRightOffset()
                            # If right offset is 0 or very small, we're at source end
                            if right_offset is not None and right_offset <= 2:
                                issues.append({
                                    'type': 'SOURCE_END',
                                    'severity': 'INFO',
                                    'start': item.GetStart(),
                                    'end': item.GetEnd(),
                                    'duration': item.GetDuration(),
                                    'track': 'V{}'.format(track_idx),
                                    'clip': item.GetName(),
                                    'message': 'Clip at source end on V{}: "{}" (right offset: {} frames)'.format(
                                        track_idx, item.GetName(), right_offset)
                                })
            except:
                pass

    return issues


def generate_report(issues, timeline_name, fps, timeline_start, timeline_end):
    """Generate a formatted QC report"""
    report = []
    report.append("")
    report.append("=" * 70)
    report.append("  TIMELINE QC REPORT")
    report.append("=" * 70)
    report.append("")
    report.append("Timeline: {}".format(timeline_name))
    report.append("Frame Rate: {} fps".format(fps))
    report.append("Duration: {} - {}".format(
        frames_to_tc(timeline_start, fps),
        frames_to_tc(timeline_end, fps)))
    report.append("")

    if not issues:
        report.append("  *** NO ISSUES FOUND ***")
        report.append("")
        report.append("=" * 70)
        return "\n".join(report)

    # Count by severity
    errors = [i for i in issues if i['severity'] == 'ERROR']
    warnings = [i for i in issues if i['severity'] == 'WARNING']
    infos = [i for i in issues if i['severity'] == 'INFO']

    report.append("SUMMARY:")
    report.append("  Errors:   {}".format(len(errors)))
    report.append("  Warnings: {}".format(len(warnings)))
    report.append("  Info:     {}".format(len(infos)))
    report.append("")

    # Sort issues by position
    issues.sort(key=lambda x: x['start'])

    # Group by type
    issue_types = {}
    for issue in issues:
        itype = issue['type']
        if itype not in issue_types:
            issue_types[itype] = []
        issue_types[itype].append(issue)

    # Print each type
    type_labels = {
        'VIDEO_GAP': 'VIDEO GAPS',
        'FLASH_FRAME': 'FLASH FRAMES',
        'AUDIO_OVERLAP': 'AUDIO OVERLAPS',
        'AUDIO_GAP': 'AUDIO GAPS',
        'DISABLED_CLIP': 'DISABLED CLIPS',
        'MUTED_TRACK': 'MUTED TRACKS',
        'OFFLINE_MEDIA': 'OFFLINE MEDIA',
        'SOURCE_END': 'CLIPS AT SOURCE END',
    }

    for itype, type_issues in issue_types.items():
        report.append("-" * 70)
        report.append("{} ({})".format(type_labels.get(itype, itype), len(type_issues)))
        report.append("-" * 70)

        for issue in type_issues:
            tc = frames_to_tc(issue['start'], fps)
            severity_marker = {
                'ERROR': '[!]',
                'WARNING': '[?]',
                'INFO': '[ ]'
            }.get(issue['severity'], '[ ]')

            report.append("  {} {} - {}".format(severity_marker, tc, issue['message']))

        report.append("")

    report.append("=" * 70)
    report.append("  END OF REPORT")
    report.append("=" * 70)
    report.append("")

    return "\n".join(report)


def main():
    print("")
    print("=" * 60)
    print("  Timeline Quality Control")
    print("=" * 60)

    resolve = get_resolve()
    if not resolve:
        print("ERROR: Could not connect to DaVinci Resolve")
        return

    project = resolve.GetProjectManager().GetCurrentProject()
    if not project:
        print("ERROR: No project open")
        return

    timeline = project.GetCurrentTimeline()
    if not timeline:
        print("ERROR: No timeline open")
        return

    fps = float(timeline.GetSetting("timelineFrameRate"))
    timeline_name = timeline.GetName()
    timeline_start = timeline.GetStartFrame()
    timeline_end = timeline.GetEndFrame()

    print("")
    print("Analyzing timeline: {}".format(timeline_name))
    print("Frame rate: {} fps".format(fps))
    print("")

    all_issues = []

    # Run all checks
    print("Checking for video gaps...")
    all_issues.extend(check_video_gaps(timeline, fps, timeline_start, timeline_end))

    print("Checking for flash frames...")
    all_issues.extend(check_flash_frames(timeline, fps))

    print("Checking for audio overlaps...")
    all_issues.extend(check_audio_overlaps(timeline, fps))

    print("Checking for audio gaps...")
    all_issues.extend(check_audio_gaps(timeline, fps))

    print("Checking for disabled/muted clips...")
    all_issues.extend(check_disabled_clips(timeline, fps))

    print("Checking for offline media...")
    all_issues.extend(check_offline_media(timeline, fps))

    print("Checking for clips at source end...")
    all_issues.extend(check_source_end(timeline, fps))

    # Generate and print report
    report = generate_report(all_issues, timeline_name, fps, timeline_start, timeline_end)
    print(report)

    # Summary
    if all_issues:
        errors = len([i for i in all_issues if i['severity'] == 'ERROR'])
        if errors > 0:
            print("ACTION REQUIRED: {} error(s) found!".format(errors))
    else:
        print("Timeline passed QC - no issues found!")


if __name__ == "__main__":
    main()
