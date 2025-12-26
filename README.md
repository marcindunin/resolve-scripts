# DaVinci Resolve Scripts

A collection of Python scripts for DaVinci Resolve automation.

## Scripts

### resolve_auto_align_multitrack.py

Automatically aligns multitrack video clips to an AAF timeline based on matching timecode.

**Use case:** When you have an AAF from audio post (Pro Tools, etc.) with audio clips, and you want to automatically place the corresponding multitrack video files at the correct timeline positions.

#### How it works

1. Reads audio clips from track 1 of the current timeline (your imported AAF)
2. Matches each audio clip's source timecode to multitrack clips in a selected bin
3. Places the matching video clips on V1 at the correct timeline positions

#### Installation

**macOS:**
```
~/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Utility/
```

**Windows:**
```
%APPDATA%\Blackmagic Design\DaVinci Resolve\Fusion\Scripts\Utility\
```

**Linux:**
```
~/.local/share/DaVinciResolve/Fusion/Scripts/Utility/
```

#### Usage

1. Import your AAF into DaVinci Resolve
2. Import your multitrack video clips into a bin (e.g., named "TRACKS")
3. Open the AAF timeline
4. Run the script from: **Workspace > Scripts > resolve_auto_align_multitrack**
5. Select the bin containing your multitrack clips from the dropdown dialog
6. The script will match and place clips automatically

#### Configuration

Edit the top of the script to customize:

```python
VIDEO_TRACK_INDEX = 1          # Target video track for placed clips
IGNORE_PREFIXES = ["Sample", "Fade"]  # Skip clips starting with these prefixes
```

#### Requirements

- DaVinci Resolve 18+ (tested on Resolve 20)
- Multitrack clips must have valid timecode matching the AAF source clips

---

### resolve_timeline_qc.py

Quality Control script that checks your timeline for common issues and generates a report.

**Use case:** Before final export, run this script to catch blank frames, audio overlaps, flash frames, and other potential issues.

#### Checks performed

- **Video gaps** - Frames with no video on any track
- **Flash frames** - Very short clips (< 3 frames by default), filters out AAF artifacts
- **Audio overlaps** - Two clips overlapping on the same audio track
- **Audio gaps** - Missing audio between clips
- **Disabled/muted clips** - Clips or tracks that may have been accidentally disabled
- **Offline media** - Verifies source files exist on disk using file system check
- **Clips at source end** - Clips trimmed to the very end of source (disabled by default)

#### Features

- Automatically skips adjustment clips in all checks
- Filters out AAF artifacts ("Sample accurate edit", "Fade" clips)
- Uses actual file system verification for offline media detection

#### Usage

1. Open the timeline you want to check
2. Run the script from: **Workspace > Scripts > resolve_timeline_qc**
3. Review the report in the console

#### Report format

```
======================================================================
  TIMELINE QC REPORT
======================================================================

Timeline: My Project v2
Frame Rate: 24.0 fps
Duration: 00:00:00:00 - 01:23:45:12

SUMMARY:
  Errors:   2
  Warnings: 5
  Info:     3

----------------------------------------------------------------------
VIDEO GAPS (2)
----------------------------------------------------------------------
  [!] 00:15:32:10 - Video gap (24 frames)
  [!] 00:45:12:03 - Video gap (12 frames)
```

#### Configuration

Edit the top of the script to customize:

```python
FLASH_FRAME_THRESHOLD = 3      # Clips shorter than this are flagged
CHECK_AUDIO_GAPS = True        # Set False if audio gaps are intentional
MIN_AUDIO_GAP_FRAMES = 2       # Ignore gaps smaller than this
IGNORE_TRACK_NAMES = []        # Track names to skip (e.g., ["Music", "SFX"])
IGNORE_ADJUSTMENT_CLIPS = True # Skip adjustment clips in all checks
IGNORE_PREFIXES = ["Sample", "Fade"]  # Skip audio clips starting with these
CHECK_OFFLINE_MEDIA = True     # Check for offline/missing media files
CHECK_SOURCE_END = False       # Check for clips at end of source media
```

#### Known limitations

- **Clip fade handles** - The Resolve scripting API does not expose clip fade handles (the fades you create by dragging clip corners). These cannot be detected by any script.

---

## Installation

Copy any script to the Resolve scripts folder:

**macOS:**
```
~/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Utility/
```

**Windows:**
```
%APPDATA%\Blackmagic Design\DaVinci Resolve\Fusion\Scripts\Utility\
```

**Linux:**
```
~/.local/share/DaVinciResolve/Fusion/Scripts/Utility/
```

Then access via **Workspace > Scripts** in DaVinci Resolve.

## License

MIT License - feel free to use and modify.
