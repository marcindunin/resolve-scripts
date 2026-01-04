#!/usr/bin/env python
# DaVinci Resolve - Timeline Quality Control
# ===========================================
# Checks timeline for common issues and generates a report.
#
# Features:
# - GUI for settings configuration
# - Real-time progress display
# - Interactive results with click-to-jump navigation
#
# Checks performed:
# - Video gaps (frames with no video content)
# - Flash frames (very short clips)
# - Audio overlaps on same track
# - Audio gaps
# - Disabled/muted clips
# - Offline media
# - Clips at end of source media

import json
import os

# ============== DEFAULT CONFIG ==============
DEFAULT_CONFIG = {
    'flash_frame_threshold': 3,
    'check_audio_gaps': True,
    'min_audio_gap_frames': 2,
    'ignore_track_names': [],
    'ignore_adjustment_clips': True,
    'ignore_prefixes': ["Sample", "Fade"],
    'check_offline_media': True,
    'check_source_end': False,
    'check_audio_overlap': True,  # Check for active audio clips overlapping across tracks
    'check_disabled_clips': True,  # Report disabled clips on video tracks
}

# Config file path (user's home directory since __file__ not available in Resolve)
def get_config_path():
    """Get config file path - works in Resolve environment"""
    try:
        # Try script directory first
        script_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(script_dir, "timeline_qc_config.json")
    except NameError:
        # __file__ not defined in Resolve - use home directory
        return os.path.join(os.path.expanduser("~"), ".timeline_qc_config.json")

CONFIG_FILE = get_config_path()

# Global state
_config = DEFAULT_CONFIG.copy()
_resolve = None
_fusion = None
_ui = None
_disp = None
_current_issues = []
_current_issue_index = 0
_timeline = None
_fps = 24.0


def load_config():
    """Load configuration from file"""
    global _config
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                saved = json.load(f)
                _config = DEFAULT_CONFIG.copy()
                _config.update(saved)
    except Exception as e:
        print("Could not load config: {}".format(e))
        _config = DEFAULT_CONFIG.copy()


def save_config():
    """Save configuration to file"""
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(_config, f, indent=2)
        return True
    except Exception as e:
        print("Could not save config: {}".format(e))
        return False


def get_resolve():
    global _resolve
    if _resolve:
        return _resolve
    try:
        import DaVinciResolveScript as dvr
        _resolve = dvr.scriptapp("Resolve")
    except ImportError:
        _resolve = bmd.scriptapp("Resolve")
    return _resolve


def get_fusion():
    global _fusion, _ui, _disp
    if _fusion:
        return _fusion, _ui, _disp
    try:
        _fusion = bmd.scriptapp("Fusion")
        _ui = _fusion.UIManager
        _disp = bmd.UIDispatcher(_ui)
    except:
        _fusion = None
        _ui = None
        _disp = None
    return _fusion, _ui, _disp


def frames_to_tc(frames, fps):
    fps = int(round(fps))
    if frames < 0:
        frames = 0
    f = frames % fps
    s = (frames // fps) % 60
    m = (frames // (fps * 60)) % 60
    h = frames // (fps * 3600)
    return "{:02d}:{:02d}:{:02d}:{:02d}".format(h, m, s, f)


def is_adjustment_clip(item):
    """Check if a timeline item is an adjustment clip"""
    if not _config.get('ignore_adjustment_clips', True):
        return False
    try:
        media_pool_item = item.GetMediaPoolItem()
        if media_pool_item is None:
            return True
        name = item.GetName()
        if name and "Adjustment Clip" in name:
            return True
    except:
        pass
    return False


def should_skip_clip(clip_name):
    """Check if clip should be skipped based on name prefix"""
    if not clip_name:
        return False
    for prefix in _config.get('ignore_prefixes', []):
        if clip_name.startswith(prefix):
            return True
    return False


def get_track_items_sorted(timeline, track_type, track_index, skip_adjustment=True):
    """Get all items on a track, sorted by start position"""
    items = timeline.GetItemListInTrack(track_type, track_index)
    if not items:
        return []

    item_list = []
    for item in items:
        if skip_adjustment and track_type == "video" and is_adjustment_clip(item):
            continue
        item_list.append({
            'item': item,
            'name': item.GetName(),
            'start': item.GetStart(),
            'end': item.GetEnd(),
            'duration': item.GetDuration(),
        })

    item_list.sort(key=lambda x: x['start'])
    return item_list


def check_video_gaps(timeline, fps, timeline_start, timeline_end, progress_callback=None):
    """Check for gaps in video coverage"""
    issues = []
    video_track_count = timeline.GetTrackCount("video")

    if video_track_count == 0:
        return issues

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

    all_video_ranges.sort(key=lambda x: x[0])

    merged = []
    for start, end in all_video_ranges:
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))

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
    threshold = _config.get('flash_frame_threshold', 3)

    video_track_count = timeline.GetTrackCount("video")
    for track_idx in range(1, video_track_count + 1):
        # Skip disabled/muted tracks
        if not is_track_enabled(timeline, "video", track_idx):
            continue

        items = get_track_items_sorted(timeline, "video", track_idx)
        for item in items:
            if item['duration'] < threshold:
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

    audio_track_count = timeline.GetTrackCount("audio")
    for track_idx in range(1, audio_track_count + 1):
        # Skip disabled/muted tracks
        if not is_track_enabled(timeline, "audio", track_idx):
            continue

        items = get_track_items_sorted(timeline, "audio", track_idx)
        for item in items:
            if should_skip_clip(item['name']):
                continue
            if item['duration'] < threshold:
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


def is_track_enabled(timeline, track_type, track_idx):
    """Check if a track is enabled (not muted)"""
    try:
        return timeline.GetIsTrackEnabled(track_type, track_idx) != False
    except:
        return True  # Assume enabled if we can't check


def is_clip_enabled(item):
    """Check if a clip is enabled (not disabled)"""
    try:
        return item.GetClipEnabled() == True
    except:
        return True  # Assume enabled if we can't check


def check_audio_overlaps(timeline, fps):
    """Check for active audio clips overlapping across different tracks"""
    if not _config.get('check_audio_overlap', True):
        return []

    issues = []
    audio_track_count = timeline.GetTrackCount("audio")

    # Collect all active audio clips with their ranges
    active_clips = []

    for track_idx in range(1, audio_track_count + 1):
        # Skip muted tracks
        if not is_track_enabled(timeline, "audio", track_idx):
            continue

        track_name = timeline.GetTrackName("audio", track_idx)
        if track_name in _config.get('ignore_track_names', []):
            continue

        items = timeline.GetItemListInTrack("audio", track_idx)
        if not items:
            continue

        for item in items:
            # Skip disabled clips
            if not is_clip_enabled(item):
                continue

            clip_name = item.GetName()

            # Skip clips with empty names (often disabled or transition artifacts)
            if not clip_name or clip_name.strip() == "":
                continue

            # Skip transition clips
            if clip_name.lower() == "transition":
                continue

            # Skip ignored prefixes
            if should_skip_clip(clip_name):
                continue

            active_clips.append({
                'item': item,
                'name': clip_name,
                'start': item.GetStart(),
                'end': item.GetEnd(),
                'track': track_idx,
            })

    # Sort by start position
    active_clips.sort(key=lambda x: x['start'])

    # Find overlaps between clips on different tracks
    reported_overlaps = set()  # Avoid duplicate reports

    for i, clip_a in enumerate(active_clips):
        for clip_b in active_clips[i+1:]:
            # Only check clips on different tracks
            if clip_a['track'] == clip_b['track']:
                continue

            # Check if they overlap in time
            # clip_b starts after clip_a (due to sorting)
            if clip_b['start'] >= clip_a['end']:
                # No overlap and no future clips will overlap with clip_a
                break

            # They overlap!
            overlap_start = clip_b['start']
            overlap_end = min(clip_a['end'], clip_b['end'])
            overlap_duration = overlap_end - overlap_start

            # Create unique key to avoid duplicates
            key = (overlap_start, clip_a['track'], clip_b['track'])
            if key in reported_overlaps:
                continue
            reported_overlaps.add(key)

            issues.append({
                'type': 'AUDIO_OVERLAP',
                'severity': 'WARNING',
                'start': overlap_start,
                'end': overlap_end,
                'duration': overlap_duration,
                'track': 'A{}/A{}'.format(clip_a['track'], clip_b['track']),
                'message': 'Audio overlap A{}/A{}: "{}" and "{}" ({} frames)'.format(
                    clip_a['track'], clip_b['track'],
                    clip_a['name'], clip_b['name'], overlap_duration)
            })

    return issues


def check_audio_gaps(timeline, fps):
    """Check for gaps in audio tracks"""
    if not _config.get('check_audio_gaps', True):
        return []

    issues = []
    audio_track_count = timeline.GetTrackCount("audio")
    min_gap = _config.get('min_audio_gap_frames', 2)
    ignore_tracks = _config.get('ignore_track_names', [])

    for track_idx in range(1, audio_track_count + 1):
        # Skip disabled/muted tracks
        if not is_track_enabled(timeline, "audio", track_idx):
            continue

        track_name = timeline.GetTrackName("audio", track_idx)
        if track_name in ignore_tracks:
            continue

        items = get_track_items_sorted(timeline, "audio", track_idx)
        if len(items) < 2:
            continue

        for i in range(len(items) - 1):
            current = items[i]
            next_item = items[i + 1]

            gap = next_item['start'] - current['end']
            if gap >= min_gap:
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

    # Check for disabled video clips (if enabled in settings)
    if _config.get('check_disabled_clips', True):
        video_track_count = timeline.GetTrackCount("video")
        for track_idx in range(1, video_track_count + 1):
            items = timeline.GetItemListInTrack("video", track_idx)
            if not items:
                continue
            for item in items:
                if is_adjustment_clip(item):
                    continue
                try:
                    if not is_clip_enabled(item):
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

    # Always check for muted audio tracks
    audio_track_count = timeline.GetTrackCount("audio")
    for track_idx in range(1, audio_track_count + 1):
        try:
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
    if not _config.get('check_offline_media', True):
        return []

    issues = []

    video_track_count = timeline.GetTrackCount("video")
    for track_idx in range(1, video_track_count + 1):
        items = timeline.GetItemListInTrack("video", track_idx)
        if not items:
            continue
        for item in items:
            if is_adjustment_clip(item):
                continue
            try:
                media_pool_item = item.GetMediaPoolItem()
                if not media_pool_item:
                    continue

                clip_props = media_pool_item.GetClipProperty()
                if not clip_props:
                    continue

                file_path = clip_props.get('File Path')
                if file_path and isinstance(file_path, str) and len(file_path) > 0:
                    if not os.path.exists(file_path):
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
    if not _config.get('check_source_end', False):
        return []

    issues = []

    video_track_count = timeline.GetTrackCount("video")
    for track_idx in range(1, video_track_count + 1):
        items = timeline.GetItemListInTrack("video", track_idx)
        if not items:
            continue
        for item in items:
            if is_adjustment_clip(item):
                continue
            try:
                media_pool_item = item.GetMediaPoolItem()
                if media_pool_item:
                    clip_props = media_pool_item.GetClipProperty()
                    if clip_props:
                        source_frames = clip_props.get('Frames')
                        if source_frames:
                            source_frames = int(source_frames)
                            right_offset = item.GetRightOffset()
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


def run_qc_analysis(timeline, progress_callback=None):
    """Run all QC checks and return issues"""
    global _fps

    _fps = float(timeline.GetSetting("timelineFrameRate"))
    timeline_start = timeline.GetStartFrame()
    timeline_end = timeline.GetEndFrame()

    all_issues = []

    if progress_callback:
        progress_callback("Checking for video gaps...")
    all_issues.extend(check_video_gaps(timeline, _fps, timeline_start, timeline_end))

    if progress_callback:
        progress_callback("Checking for flash frames...")
    all_issues.extend(check_flash_frames(timeline, _fps))

    if progress_callback:
        progress_callback("Checking for audio overlaps...")
    all_issues.extend(check_audio_overlaps(timeline, _fps))

    if progress_callback:
        progress_callback("Checking for audio gaps...")
    all_issues.extend(check_audio_gaps(timeline, _fps))

    if progress_callback:
        progress_callback("Checking for disabled/muted clips...")
    all_issues.extend(check_disabled_clips(timeline, _fps))

    if progress_callback:
        progress_callback("Checking for offline media...")
    all_issues.extend(check_offline_media(timeline, _fps))

    if progress_callback:
        progress_callback("Checking for clips at source end...")
    all_issues.extend(check_source_end(timeline, _fps))

    # Sort by position
    all_issues.sort(key=lambda x: x['start'])

    return all_issues


def jump_to_timecode(timeline, frame):
    """Jump to a specific frame in the timeline"""
    try:
        timeline.SetCurrentTimecode(frames_to_tc(frame, _fps))
    except:
        try:
            # Alternative method
            resolve = get_resolve()
            resolve.OpenPage("edit")
            timeline.SetCurrentTimecode(frames_to_tc(frame, _fps))
        except Exception as e:
            print("Could not jump to timecode: {}".format(e))


# ============== GUI WINDOWS ==============

def show_settings_window():
    """Show the settings configuration window"""
    global _config

    fusion, ui, disp = get_fusion()
    if not fusion:
        print("ERROR: Could not get Fusion UI")
        return None

    load_config()

    win = disp.AddWindow({
        'ID': 'SettingsWin',
        'WindowTitle': 'Timeline QC - Settings',
        'Geometry': [100, 100, 450, 400],
        'Spacing': 10,
    }, [
        ui.VGroup({'Spacing': 5}, [
            ui.Label({'Text': 'Timeline QC Settings', 'Font': ui.Font({'PixelSize': 18, 'Bold': True}), 'Weight': 0}),
            ui.Label({'Text': '─' * 50, 'Weight': 0}),

            # Flash frame threshold
            ui.HGroup({'Weight': 0}, [
                ui.Label({'Text': 'Flash frame threshold (frames):', 'Weight': 2}),
                ui.SpinBox({'ID': 'FlashThreshold', 'Value': _config.get('flash_frame_threshold', 3), 'Minimum': 1, 'Maximum': 30, 'Weight': 1}),
            ]),

            # Min audio gap
            ui.HGroup({'Weight': 0}, [
                ui.Label({'Text': 'Min audio gap to report (frames):', 'Weight': 2}),
                ui.SpinBox({'ID': 'MinAudioGap', 'Value': _config.get('min_audio_gap_frames', 2), 'Minimum': 1, 'Maximum': 100, 'Weight': 1}),
            ]),

            # Ignore prefixes
            ui.HGroup({'Weight': 0}, [
                ui.Label({'Text': 'Ignore clip prefixes (comma-separated):', 'Weight': 2}),
                ui.LineEdit({'ID': 'IgnorePrefixes', 'Text': ', '.join(_config.get('ignore_prefixes', [])), 'Weight': 2}),
            ]),

            ui.Label({'Text': '─' * 50, 'Weight': 0}),
            ui.Label({'Text': 'Checks to perform:', 'Weight': 0}),

            # Checkboxes
            ui.CheckBox({'ID': 'CheckAudioOverlap', 'Text': 'Check audio overlaps (active clips on different tracks)', 'Checked': _config.get('check_audio_overlap', True), 'Weight': 0}),
            ui.CheckBox({'ID': 'CheckAudioGaps', 'Text': 'Check audio gaps', 'Checked': _config.get('check_audio_gaps', True), 'Weight': 0}),
            ui.CheckBox({'ID': 'CheckOfflineMedia', 'Text': 'Check offline media', 'Checked': _config.get('check_offline_media', True), 'Weight': 0}),
            ui.CheckBox({'ID': 'CheckDisabledClips', 'Text': 'Report disabled video clips', 'Checked': _config.get('check_disabled_clips', True), 'Weight': 0}),
            ui.CheckBox({'ID': 'CheckSourceEnd', 'Text': 'Check clips at source end', 'Checked': _config.get('check_source_end', False), 'Weight': 0}),
            ui.CheckBox({'ID': 'IgnoreAdjustment', 'Text': 'Ignore adjustment clips', 'Checked': _config.get('ignore_adjustment_clips', True), 'Weight': 0}),

            ui.Label({'Text': '', 'Weight': 1}),  # Spacer

            ui.Label({'Text': '─' * 50, 'Weight': 0}),

            # Buttons
            ui.HGroup({'Weight': 0}, [
                ui.Button({'ID': 'SaveBtn', 'Text': 'Save Settings', 'Weight': 1}),
                ui.Button({'ID': 'StartBtn', 'Text': 'Start QC', 'Weight': 1}),
                ui.Button({'ID': 'CancelBtn', 'Text': 'Cancel', 'Weight': 1}),
            ]),
        ]),
    ])

    result = {'action': None}

    def on_save(ev):
        _config['flash_frame_threshold'] = win.Find('FlashThreshold').Value
        _config['min_audio_gap_frames'] = win.Find('MinAudioGap').Value
        _config['check_audio_overlap'] = win.Find('CheckAudioOverlap').Checked
        _config['check_audio_gaps'] = win.Find('CheckAudioGaps').Checked
        _config['check_offline_media'] = win.Find('CheckOfflineMedia').Checked
        _config['check_disabled_clips'] = win.Find('CheckDisabledClips').Checked
        _config['check_source_end'] = win.Find('CheckSourceEnd').Checked
        _config['ignore_adjustment_clips'] = win.Find('IgnoreAdjustment').Checked

        prefixes_text = win.Find('IgnorePrefixes').Text
        _config['ignore_prefixes'] = [p.strip() for p in prefixes_text.split(',') if p.strip()]

        if save_config():
            print("Settings saved!")

    def on_start(ev):
        on_save(ev)  # Save settings first
        result['action'] = 'start'
        disp.ExitLoop()

    def on_cancel(ev):
        result['action'] = 'cancel'
        disp.ExitLoop()

    def on_close(ev):
        result['action'] = 'cancel'
        disp.ExitLoop()

    win.On.SaveBtn.Clicked = on_save
    win.On.StartBtn.Clicked = on_start
    win.On.CancelBtn.Clicked = on_cancel
    win.On.SettingsWin.Close = on_close

    win.Show()
    disp.RunLoop()
    win.Hide()

    return result['action']


def show_progress_window(timeline):
    """Show progress window and run analysis"""
    global _current_issues, _timeline

    _timeline = timeline

    fusion, ui, disp = get_fusion()
    if not fusion:
        return []

    win = disp.AddWindow({
        'ID': 'ProgressWin',
        'WindowTitle': 'Timeline QC - Analyzing...',
        'Geometry': [100, 100, 500, 300],
        'Spacing': 10,
    }, [
        ui.VGroup({'Spacing': 5}, [
            ui.Label({'Text': 'Analyzing Timeline...', 'Font': ui.Font({'PixelSize': 16, 'Bold': True}), 'Weight': 0, 'ID': 'StatusLabel'}),
            ui.Label({'Text': 'Timeline: ' + timeline.GetName(), 'Weight': 0}),
            ui.Label({'Text': '─' * 60, 'Weight': 0}),
            ui.TextEdit({'ID': 'ConsoleOutput', 'ReadOnly': True, 'Font': ui.Font({'Family': 'Consolas', 'PixelSize': 12}), 'Weight': 1}),
            ui.Label({'Text': '─' * 60, 'Weight': 0}),
            ui.Button({'ID': 'CloseBtn', 'Text': 'Please wait...', 'Enabled': False, 'Weight': 0}),
        ]),
    ])

    console = win.Find('ConsoleOutput')
    close_btn = win.Find('CloseBtn')
    status_label = win.Find('StatusLabel')

    output_lines = []

    def add_output(text):
        output_lines.append(text)
        console.PlainText = '\n'.join(output_lines)

    def on_close(ev):
        disp.ExitLoop()

    win.On.CloseBtn.Clicked = on_close
    win.On.ProgressWin.Close = on_close

    win.Show()

    # Run analysis
    add_output("Starting QC analysis...")
    add_output("Frame rate: {} fps".format(timeline.GetSetting("timelineFrameRate")))
    add_output("")

    # We need to run analysis in steps to update UI
    # Since Resolve scripting is single-threaded, we do it sequentially
    _current_issues = run_qc_analysis(timeline, add_output)

    add_output("")
    add_output("=" * 50)
    add_output("Analysis complete!")
    add_output("")

    errors = len([i for i in _current_issues if i['severity'] == 'ERROR'])
    warnings = len([i for i in _current_issues if i['severity'] == 'WARNING'])
    infos = len([i for i in _current_issues if i['severity'] == 'INFO'])

    add_output("Results: {} errors, {} warnings, {} info".format(errors, warnings, infos))
    add_output("Total issues: {}".format(len(_current_issues)))

    status_label.Text = "Analysis Complete!"
    close_btn.Text = "View Results"
    close_btn.Enabled = True

    disp.RunLoop()
    win.Hide()

    return _current_issues


def show_results_window(issues, timeline):
    """Show results window with issue list and navigation"""
    global _current_issue_index, _timeline, _fps

    _timeline = timeline
    _fps = float(timeline.GetSetting("timelineFrameRate"))

    fusion, ui, disp = get_fusion()
    if not fusion:
        return

    _current_issue_index = 0

    # Prepare issue list for display
    issue_rows = []
    for i, issue in enumerate(issues):
        tc = frames_to_tc(issue['start'], _fps)
        severity = issue['severity']
        message = issue['message']
        issue_rows.append({
            'index': i,
            'tc': tc,
            'severity': severity,
            'message': message,
            'frame': issue['start']
        })

    # Summary
    errors = len([i for i in issues if i['severity'] == 'ERROR'])
    warnings = len([i for i in issues if i['severity'] == 'WARNING'])
    infos = len([i for i in issues if i['severity'] == 'INFO'])

    if not issues:
        summary_text = "No issues found! Timeline passed QC."
        summary_color = {'R': 0.2, 'G': 0.8, 'B': 0.2, 'A': 1}
    elif errors > 0:
        summary_text = "{} errors, {} warnings, {} info".format(errors, warnings, infos)
        summary_color = {'R': 1, 'G': 0.3, 'B': 0.3, 'A': 1}
    else:
        summary_text = "{} warnings, {} info".format(warnings, infos)
        summary_color = {'R': 1, 'G': 0.8, 'B': 0.2, 'A': 1}

    win = disp.AddWindow({
        'ID': 'ResultsWin',
        'WindowTitle': 'Timeline QC - Results',
        'Geometry': [100, 100, 700, 500],
        'Spacing': 10,
    }, [
        ui.VGroup({'Spacing': 5}, [
            ui.Label({'Text': 'QC Results: ' + timeline.GetName(), 'Font': ui.Font({'PixelSize': 16, 'Bold': True}), 'Weight': 0}),
            ui.Label({'Text': summary_text, 'Weight': 0, 'ID': 'SummaryLabel'}),
            ui.Label({'Text': '─' * 80, 'Weight': 0}),

            # Issue list
            ui.Tree({
                'ID': 'IssueTree',
                'Weight': 1,
                'HeaderHidden': False,
                'SelectionMode': 'SingleSelection',
                'Events': {'ItemClicked': True, 'ItemDoubleClicked': True},
            }),

            ui.Label({'Text': '─' * 80, 'Weight': 0}),

            # Navigation
            ui.HGroup({'Weight': 0, 'Spacing': 10}, [
                ui.Label({'Text': 'Navigate:', 'Weight': 0}),
                ui.Button({'ID': 'PrevBtn', 'Text': '< Previous', 'Weight': 1}),
                ui.Label({'ID': 'PositionLabel', 'Text': '0 / 0', 'Alignment': {'AlignHCenter': True}, 'Weight': 1}),
                ui.Button({'ID': 'NextBtn', 'Text': 'Next >', 'Weight': 1}),
                ui.Button({'ID': 'JumpBtn', 'Text': 'Jump to Selected', 'Weight': 1}),
            ]),

            ui.HGroup({'Weight': 0}, [
                ui.Button({'ID': 'ExportBtn', 'Text': 'Export Report', 'Weight': 1}),
                ui.Button({'ID': 'CloseBtn', 'Text': 'Close', 'Weight': 1}),
            ]),
        ]),
    ])

    # Setup tree
    tree = win.Find('IssueTree')
    tree_header = tree.NewItem()
    tree_header.Text[0] = '#'
    tree_header.Text[1] = 'Timecode'
    tree_header.Text[2] = 'Severity'
    tree_header.Text[3] = 'Issue'
    tree.SetHeaderItem(tree_header)
    tree.ColumnCount = 4
    tree.ColumnWidth[0] = 40
    tree.ColumnWidth[1] = 100
    tree.ColumnWidth[2] = 80
    tree.ColumnWidth[3] = 450

    # Populate tree
    tree_items = []
    for row in issue_rows:
        item = tree.NewItem()
        item.Text[0] = str(row['index'] + 1)
        item.Text[1] = row['tc']
        item.Text[2] = row['severity']
        item.Text[3] = row['message']
        tree.AddTopLevelItem(item)
        tree_items.append(item)

    position_label = win.Find('PositionLabel')

    def update_position_label():
        if issues:
            position_label.Text = "{} / {}".format(_current_issue_index + 1, len(issues))
        else:
            position_label.Text = "0 / 0"

    def select_current_issue():
        if issues and 0 <= _current_issue_index < len(tree_items):
            try:
                tree_items[_current_issue_index].Selected = True
            except:
                pass

    def jump_to_current():
        if issues and 0 <= _current_issue_index < len(issues):
            frame = issues[_current_issue_index]['start']
            jump_to_timecode(_timeline, frame)

    def on_prev(ev):
        global _current_issue_index
        if issues and _current_issue_index > 0:
            _current_issue_index -= 1
            update_position_label()
            select_current_issue()
            jump_to_current()

    def on_next(ev):
        global _current_issue_index
        if issues and _current_issue_index < len(issues) - 1:
            _current_issue_index += 1
            update_position_label()
            select_current_issue()
            jump_to_current()

    def on_jump(ev):
        jump_to_current()

    def on_item_clicked(ev):
        global _current_issue_index
        item = ev.get('item')
        if item:
            try:
                idx = int(item.Text[0]) - 1
                if 0 <= idx < len(issues):
                    _current_issue_index = idx
                    update_position_label()
            except:
                pass

    def on_item_double_clicked(ev):
        global _current_issue_index
        item = ev.get('item')
        if item:
            try:
                idx = int(item.Text[0]) - 1
                if 0 <= idx < len(issues):
                    _current_issue_index = idx
                    update_position_label()
                    jump_to_current()
            except:
                pass

    def on_export(ev):
        # Generate text report
        report_lines = []
        report_lines.append("TIMELINE QC REPORT")
        report_lines.append("=" * 60)
        report_lines.append("Timeline: {}".format(timeline.GetName()))
        report_lines.append("Frame Rate: {} fps".format(_fps))
        report_lines.append("")
        report_lines.append("SUMMARY: {} errors, {} warnings, {} info".format(errors, warnings, infos))
        report_lines.append("")
        report_lines.append("-" * 60)

        for issue in issues:
            tc = frames_to_tc(issue['start'], _fps)
            report_lines.append("[{}] {} - {}".format(issue['severity'], tc, issue['message']))

        report_lines.append("-" * 60)
        report_lines.append("END OF REPORT")

        report_text = '\n'.join(report_lines)

        # Create safe default filename
        safe_name = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in timeline.GetName())
        default_name = "QC_Report_{}.txt".format(safe_name.replace(" ", "_"))

        # Get Desktop path - try different methods
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.exists(desktop_path):
            desktop_path = os.path.expanduser("~")

        default_path = os.path.join(desktop_path, default_name)

        # Show custom save dialog
        save_result = {'path': None}

        save_win = disp.AddWindow({
            'ID': 'SaveDialog',
            'WindowTitle': 'Save QC Report',
            'Geometry': [200, 200, 500, 150],
            'Spacing': 10,
        }, [
            ui.VGroup({'Spacing': 5}, [
                ui.Label({'Text': 'Save report to:', 'Weight': 0}),
                ui.HGroup({'Weight': 0}, [
                    ui.LineEdit({'ID': 'SavePath', 'Text': default_path, 'Weight': 3}),
                    ui.Button({'ID': 'BrowseBtn', 'Text': 'Browse...', 'Weight': 1}),
                ]),
                ui.HGroup({'Weight': 0}, [
                    ui.Button({'ID': 'SaveBtn', 'Text': 'Save', 'Weight': 1}),
                    ui.Button({'ID': 'CancelSaveBtn', 'Text': 'Cancel', 'Weight': 1}),
                ]),
            ]),
        ])

        def on_browse(ev):
            # Use RequestDir to select folder
            try:
                folder = fusion.RequestDir(desktop_path)
                if folder:
                    current_name = os.path.basename(save_win.Find('SavePath').Text)
                    if not current_name:
                        current_name = default_name
                    save_win.Find('SavePath').Text = os.path.join(folder, current_name)
            except:
                pass

        def on_save_file(ev):
            save_result['path'] = save_win.Find('SavePath').Text
            disp.ExitLoop()

        def on_cancel_save(ev):
            save_result['path'] = None
            disp.ExitLoop()

        def on_close_save(ev):
            save_result['path'] = None
            disp.ExitLoop()

        save_win.On.BrowseBtn.Clicked = on_browse
        save_win.On.SaveBtn.Clicked = on_save_file
        save_win.On.CancelSaveBtn.Clicked = on_cancel_save
        save_win.On.SaveDialog.Close = on_close_save

        save_win.Show()
        disp.RunLoop()
        save_win.Hide()

        if save_result['path']:
            report_path = save_result['path']
            # Ensure .txt extension
            if not report_path.lower().endswith('.txt'):
                report_path += '.txt'

            try:
                # Create directory if needed
                dir_path = os.path.dirname(report_path)
                if dir_path and not os.path.exists(dir_path):
                    os.makedirs(dir_path)

                with open(report_path, 'w', encoding='utf-8') as f:
                    f.write(report_text)
                print("=" * 50)
                print("Report saved to:")
                print(report_path)
                print("=" * 50)
            except Exception as e:
                print("Could not save report: {}".format(e))
        else:
            print("Export cancelled")

    def on_close(ev):
        disp.ExitLoop()

    win.On.PrevBtn.Clicked = on_prev
    win.On.NextBtn.Clicked = on_next
    win.On.JumpBtn.Clicked = on_jump
    win.On.IssueTree.ItemClicked = on_item_clicked
    win.On.IssueTree.ItemDoubleClicked = on_item_double_clicked
    win.On.ExportBtn.Clicked = on_export
    win.On.CloseBtn.Clicked = on_close
    win.On.ResultsWin.Close = on_close

    update_position_label()
    if tree_items:
        try:
            tree_items[0].Selected = True
        except:
            pass

    win.Show()
    disp.RunLoop()
    win.Hide()


def main():
    """Main entry point"""
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

    fusion, ui, disp = get_fusion()
    if not fusion:
        print("ERROR: Could not get Fusion UI - running in console mode")
        # Fallback to console mode
        load_config()
        issues = run_qc_analysis(timeline, print)
        print("\nFound {} issues".format(len(issues)))
        for issue in issues:
            print("[{}] {} - {}".format(
                issue['severity'],
                frames_to_tc(issue['start'], _fps),
                issue['message']
            ))
        return

    # Show settings window
    action = show_settings_window()

    if action != 'start':
        print("Cancelled")
        return

    # Show progress and run analysis
    issues = show_progress_window(timeline)

    # Show results
    show_results_window(issues, timeline)

    print("")
    print("QC Complete!")


if __name__ == "__main__":
    main()
