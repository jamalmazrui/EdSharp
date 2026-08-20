---
title: "VLC Help Guide"
subtitle: "VideoLAN's official VLC media player documentation, gathered into one document"
date: 2026-08-02
lang: en
---

# VLC Help Guide

This guide gathers VideoLAN's official VLC documentation, as it stood on 2 August 2026.

**A word about where this came from, because it matters.** The VideoLAN wiki as a whole is
community writing, and VideoLAN says so: its stated purpose is to document what is not in
the official documentation. But inside that wiki sits a separate section, the Documentation
namespace, and its pages carry VideoLAN's own banner declaring them part of the official VLC
media player Documentation. Only that namespace was gathered. Of the 316 pages taken, 278
carry the banner.

**This documentation is old, and you should know that before relying on it.** VideoLAN's
User Guide pages carry copyright dates running to 2009; the Quick Start Guide was last
revised in 2019. VideoLAN says of its own Modules documentation that parts of it are
currently outdated or incomplete. VLC is now at version 3 and heading to 4. The
fundamentals here -- playing, converting, streaming, the command line -- have changed
little. Descriptions of menus and buttons may not match what you find.

**Two thirds of this guide is the module and option reference.** VLC is built from modules,
and VideoLAN documents each one with its options. Those are short, factual and genuinely
useful from the command line, and they are the part VideoLAN itself flags as patchy. They
are kept, in their own category, so the rest of the guide is not diluted by them.

What was left out, and why:

- **The Hacker Guide and the developer pages** -- writing modules, the source tree, module
  loading, Lua scripting, building HTTP interface pages -- under this collection's standing
  exclusion of developer documentation.
- **The rest of the VideoLAN wiki**, which is community writing rather than VideoLAN's own.
- **The wiki's editing guidelines and its own index pages**, which are housekeeping.
- **Translations.** One language per volume.

Cross-references between pages have been resolved into links within this document, so
following one never leaves the file. Illustrations do not survive the conversion.

This document holds 9 categories, 280 articles and about 125,867 words.

## Contents {#contents}

- [Audio, Video and Subtitles](#audio-video-and-subtitles)
    - [Audio](#audio)
    - [Subtitles](#subtitles)
    - [Video](#video)
    - [Video and Audio Filters](#video-and-audio-filters)
- [Getting Started](#getting-started)
    - [Install VLC](#install-vlc)
    - [Installing VLC](#installing-vlc)
    - [Quick Start Guide](#quick-start-guide)
    - [Uninstalling VLC](#uninstalling-vlc)
    - [User Guide](#user-guide)
    - [VLC for dummies](#vlc-for-dummies)
- [Modules and Options](#modules-and-options)
    - [a52](#modules-a52)
    - [adjust](#modules-adjust)
    - [adpcm](#modules-adpcm)
    - [alphamask](#modules-alphamask)
    - [alsa](#modules-alsa)
    - [anaglyph](#modules-anaglyph)
    - [arts](#modules-arts)
    - [asf](#modules-asf)
    - [atmo](#modules-atmo)
    - [audioqueue](#modules-audioqueue)
    - [auhal](#modules-auhal)
    - [autodel](#modules-autodel)
    - [avcapture](#modules-avcapture)
    - [avcodec](#modules-avcodec)
    - [avformat](#modules-avformat)
    - [avi](#modules-avi)
    - [avio](#modules-avio)
    - [bandwidth](#modules-bandwidth)
    - [bda](#modules-bda)
    - [beos](#modules-beos)
    - [blend](#modules-blend)
    - [bluescreen](#modules-bluescreen)
    - [bonjour](#modules-bonjour)
    - [bridge / in](#modules-bridge-in)
    - [bridge / out](#modules-bridge-out)
    - [caca](#modules-caca)
    - [cdda](#modules-cdda)
    - [clone](#modules-clone)
    - [colorthres](#modules-colorthres)
    - [crop](#modules-crop)
    - [croppadd](#modules-croppadd)
    - [daala](#modules-daala)
    - [daap](#modules-daap)
    - [dc1394](#modules-dc1394)
    - [deinterlace](#modules-deinterlace)
    - [delay](#modules-delay)
    - [description](#modules-description)
    - [dirac](#modules-dirac)
    - [direct3d](#modules-direct3d)
    - [directfb](#modules-directfb)
    - [directory](#modules-directory)
    - [directx aout](#modules-directx-aout)
    - [directx vout](#modules-directx-vout)
    - [display](#modules-display)
    - [distort](#modules-distort)
    - [dshow](#modules-dshow)
    - [dtv](#modules-dtv)
    - [dtv / Linux options](#modules-dtv-linux-options)
    - [dtv / Windows options](#modules-dtv-windows-options)
    - [dummy](#modules-dummy)
    - [dummy sout](#modules-dummy-sout)
    - [dump](#modules-dump)
    - [duplicate](#modules-duplicate)
    - [dvb](#modules-dvb)
    - [dvbsub](#modules-dvbsub)
    - [dvdnav](#modules-dvdnav)
    - [dvdread](#modules-dvdread)
    - [erase](#modules-erase)
    - [esd](#modules-esd)
    - [extract](#modules-extract)
    - [eyetv](#modules-eyetv)
    - [faad](#modules-faad)
    - [fake](#modules-fake)
    - [fdkaac](#modules-fdkaac)
    - [file](#modules-file)
    - [file aout](#modules-file-aout)
    - [flac](#modules-flac)
    - [fluidsynth](#modules-fluidsynth)
    - [freeze](#modules-freeze)
    - [ftp](#modules-ftp)
    - [galaktos](#modules-galaktos)
    - [gather](#modules-gather)
    - [gaussianblur](#modules-gaussianblur)
    - [glwin32](#modules-glwin32)
    - [glx](#modules-glx)
    - [gme](#modules-gme)
    - [goom](#modules-goom)
    - [gradfun](#modules-gradfun)
    - [gradient](#modules-gradient)
    - [h26x](#modules-h26x)
    - [hal](#modules-hal)
    - [hotkeys](#modules-hotkeys)
    - [http](#modules-http)
    - [http intf](#modules-http-intf)
    - [image](#modules-image)
    - [invert](#modules-invert)
    - [jack](#modules-jack)
    - [jpeg](#modules-jpeg)
    - [kate](#modules-kate)
    - [lirc](#modules-lirc)
    - [live](#modules-live)
    - [live555](#modules-live555)
    - [logo](#modules-logo)
    - [lua](#modules-lua)
    - [macos](#modules-macos)
    - [macosx gui](#modules-macosx-gui)
    - [magnify](#modules-magnify)
    - [marq](#modules-marq)
    - [mjpeg](#modules-mjpeg)
    - [mkv](#modules-mkv)
    - [mmdevice](#modules-mmdevice)
    - [mms](#modules-mms)
    - [mod](#modules-mod)
    - [Modules](#modules)
    - [mosaic](#modules-mosaic)
    - [mosaic / bridge](#modules-mosaic-bridge)
    - [motion control](#modules-motion-control)
    - [motionblur](#modules-motionblur)
    - [mp4](#modules-mp4)
    - [mpc](#modules-mpc)
    - [mpjpeg](#modules-mpjpeg)
    - [mqtt](#modules-mqtt)
    - [ncurses](#modules-ncurses)
    - [netsync](#modules-netsync)
    - [noise](#modules-noise)
    - [nsv](#modules-nsv)
    - [ogg](#modules-ogg)
    - [oldmovie](#modules-oldmovie)
    - [opengl](#modules-opengl)
    - [opensles](#modules-opensles)
    - [osc](#modules-osc)
    - [oss](#modules-oss)
    - [panoramix](#modules-panoramix)
    - [playlist](#modules-playlist)
    - [podcast](#modules-podcast)
    - [podcast sd](#modules-podcast-sd)
    - [portaudio](#modules-portaudio)
    - [posterize](#modules-posterize)
    - [projectm](#modules-projectm)
    - [psychedelic](#modules-psychedelic)
    - [pulse](#modules-pulse)
    - [puzzle](#modules-puzzle)
    - [pva](#modules-pva)
    - [pvr](#modules-pvr)
    - [Qt4](#modules-qt4)
    - [qtcapture](#modules-qtcapture)
    - [qtsound](#modules-qtsound)
    - [rawdv](#modules-rawdv)
    - [rawvid](#modules-rawvid)
    - [real](#modules-real)
    - [record](#modules-record)
    - [ripple](#modules-ripple)
    - [rotate](#modules-rotate)
    - [rss](#modules-rss)
    - [rtp](#modules-rtp)
    - [rtsp](#modules-rtsp)
    - [sap](#modules-sap)
    - [scale](#modules-scale)
    - [scene](#modules-scene)
    - [schroedinger](#modules-schroedinger)
    - [screen](#modules-screen)
    - [sdl aout](#modules-sdl-aout)
    - [sdl vout](#modules-sdl-vout)
    - [sdp](#modules-sdp)
    - [sepia](#modules-sepia)
    - [sharpen](#modules-sharpen)
    - [shout](#modules-shout)
    - [skins2](#modules-skins2)
    - [smem](#modules-smem)
    - [speex](#modules-speex)
    - [standard](#modules-standard)
    - [std](#modules-std)
    - [subsdelay](#modules-subsdelay)
    - [subtitle](#modules-subtitle)
    - [svg](#modules-svg)
    - [switcher](#modules-switcher)
    - [telnet](#modules-telnet)
    - [telx](#modules-telx)
    - [time](#modules-time)
    - [timeshift](#modules-timeshift)
    - [transcode](#modules-transcode)
    - [transform](#modules-transform)
    - [transrate](#modules-transrate)
    - [udp](#modules-udp)
    - [upnp](#modules-upnp)
    - [upnp cc](#modules-upnp-cc)
    - [upnp intel](#modules-upnp-intel)
    - [v4l](#modules-v4l)
    - [v4l2](#modules-v4l2)
    - [vcd](#modules-vcd)
    - [VHS](#modules-vhs)
    - [visual](#modules-visual)
    - [vobsub](#modules-vobsub)
    - [vorbis](#modules-vorbis)
    - [vpx](#modules-vpx)
    - [vsxu](#modules-vsxu)
    - [wall](#modules-wall)
    - [wav](#modules-wav)
    - [wave](#modules-wave)
    - [waveout](#modules-waveout)
    - [wiimote](#modules-wiimote)
    - [wingdi](#modules-wingdi)
    - [wxWidgets](#modules-wxwidgets)
    - [x11](#modules-x11)
    - [x264](#modules-x264)
    - [x265](#modules-x265)
    - [xvideo](#modules-xvideo)
- [Playing Media](#playing-media)
    - [Advanced Use of VLC](#play-howto-advanced-use-of-vlc)
    - [Basic Use](#play-howto-basic-use)
    - [Basic Use 0.8](#play-howto-basic-use-0-8)
    - [Basic Use 0.9](#play-howto-basic-use-0-9)
    - [Basic Use 0.9 / Audio](#play-howto-basic-use-0-9-audio)
    - [Basic Use 0.9 / Basic troubleshooting](#play-howto-basic-use-0-9-basic-troubleshooting)
    - [Basic Use 0.9 / Hotkeys](#play-howto-basic-use-0-9-hotkeys)
    - [Basic Use 0.9 / Interface](#play-howto-basic-use-0-9-interface)
    - [Basic Use 0.9 / Opening modes](#play-howto-basic-use-0-9-opening-modes)
    - [Basic Use 0.9 / Playback](#play-howto-basic-use-0-9-playback)
    - [Basic Use 0.9 / Playlist](#play-howto-basic-use-0-9-playlist)
    - [Basic Use 0.9 / Snapshots](#play-howto-basic-use-0-9-snapshots)
    - [Basic Use 0.9 / Subtitles](#play-howto-basic-use-0-9-subtitles)
    - [Basic Use 0.9 / Video](#play-howto-basic-use-0-9-video)
    - [Basic Use 0.9 / Video and Audio Filters](#play-howto-basic-use-0-9-video-and-audio-filters)
    - [Basic Use / Audio](#play-howto-basic-use-audio)
    - [Basic Use / Basic troubleshooting](#play-howto-basic-use-basic-troubleshooting)
    - [Basic Use / Hotkeys](#play-howto-basic-use-hotkeys)
    - [Basic Use / Interface](#play-howto-basic-use-interface)
    - [Basic Use / Interface in Windows 7](#play-howto-basic-use-interface-in-windows-7)
    - [Basic Use / Interface OSX](#play-howto-basic-use-interface-osx)
    - [Basic Use / Interface Windows](#play-howto-basic-use-interface-windows)
    - [Basic Use / Interface Windows 7](#play-howto-basic-use-interface-windows-7)
    - [Basic Use / Menus](#play-howto-basic-use-menus)
    - [Basic Use / Open](#play-howto-basic-use-open)
    - [Basic Use / Playback](#play-howto-basic-use-playback)
    - [Basic Use / Playlist](#play-howto-basic-use-playlist)
    - [Basic Use / Snapshots](#play-howto-basic-use-snapshots)
    - [Basic Use / Subtitles](#play-howto-basic-use-subtitles)
    - [Basic Use / Video](#play-howto-basic-use-video)
    - [Basic Use / Video and Audio Filters](#play-howto-basic-use-video-and-audio-filters)
    - [Basic Use / VLC 1.2 Interface on Ubuntu](#play-howto-basic-use-vlc-1-2-interface-on-ubuntu)
    - [Basic Use / VLC 1.2 Interface on Windows 7](#play-howto-basic-use-vlc-1-2-interface-on-windows-7)
    - [Building Lua Playlist Scripts](#play-howto-building-lua-playlist-scripts)
    - [Building Pages for the HTTP Interface](#play-howto-building-pages-for-the-http-interface)
    - [Format String](#play-howto-format-string)
    - [Introduction to VLC](#play-howto-introduction-to-vlc)
    - [Play HowTo](#play-howto)
- [Streaming and Converting](#streaming-and-converting)
    - [Advanced Streaming Using the Command Line](#streaming-howto-advanced-streaming-using-the-command-line)
    - [Advanced streaming with samples /  multiple files streaming /  using multicast in streaming](#streaming-howto-advanced-streaming-with-samples-multiple-files-streaming-using-multicast-in-streaming)
    - [Command Line Examples](#streaming-howto-command-line-examples)
    - [Easy Streaming](#streaming-howto-easy-streaming)
    - [Easy Streaming Newer Versions](#streaming-howto-easy-streaming-newer-versions)
    - [Receive and Save a Stream](#streaming-howto-receive-and-save-a-stream)
    - [Stream a DVB Channel](#streaming-howto-stream-a-dvb-channel)
    - [Stream a DVD](#streaming-howto-stream-a-dvd)
    - [Stream a File](#streaming-howto-stream-a-file)
    - [Stream from a DV Camcorder](#streaming-howto-stream-from-a-dv-camcorder)
    - [Stream from Encoding Cards and Other Capture Devices](#streaming-howto-stream-from-encoding-cards-and-other-capture-devices)
    - [Streaming /  Muxers and Codecs](#streaming-howto-streaming-muxers-and-codecs)
    - [Streaming a live video feed to Darwin Streaming Server for Mobile Phones](#streaming-howto-streaming-a-live-video-feed-to-darwin-streaming-server-for-mobile-phones)
    - [Streaming for the iPhone](#streaming-howto-streaming-for-the-iphone)
    - [Streaming HowTo](#streaming-howto)
    - [Streaming over IPv6](#streaming-howto-streaming-over-ipv6)
    - [VLM](#streaming-howto-vlm)
- [Troubleshooting](#troubleshooting)
    - [Basic troubleshooting](#basic-troubleshooting)
    - [Misc](#misc)
- [Using VLC](#using-vlc)
    - [Alternative Interfaces](#alternative-interfaces)
    - [Command line](#command-line)
    - [Hotkeys](#hotkeys)
    - [Interface](#interface)
    - [Open Media](#open-media)
    - [Playback](#playback)
    - [Playlist](#playlist)
    - [Snapshots](#snapshots)
    - [WebPlugin](#webplugin)
- [VLC on Phones and Tablets](#vlc-on-phones-and-tablets)
    - [Android](#android)
    - [IOS](#ios)
    - [Ubuntu Phone](#ubuntu-phone)
- [Miscellaneous](#miscellaneous)
    - [History](#history)

## Audio, Video and Subtitles {#audio-video-and-subtitles}

### Audio {#audio}

VLC can play several audio formats: *.asf, .avi, .divx, .dv, .mxf, .ogg, .gm, .ps, .ts, .vob,* and *.wmv*. It can convert audio tracks and use several visualizations.

**Note:** The commands in the **Audio** menu are only enabled when an audio file is being played.

#### Playing an audio track

To play a track:

1.  Select *Open File* in the *Media* menu.
2.  Select an audio file and click on the  *Play* button. The selected track is played.

#### Enabling and disabling audio tracks

-   To disable a track, select the *Disable* option in the *Audio Track* from the *Audio* menu. The selected track will then stop.
-   To enable the track again, select the designated *Track* option in the *Audio Track* from the *Audio* menu. The selected track will then play.

#### Recording Audio

To record audio you need the record button () to be visible. The record button is hidden by default. You can display using one of these methods:

-   Select Advanced Controls in the View menu. The Advanced toolbar is displayed on top of the standard toolbar. The Advanced toolbar contains the Record button.
-   Select Customize interface in the Tools menu and add the record button to the Line 2 of buttons (which is the line shown by default).

Once the Record button is visible, click it to start recording.

The recording from a shoutcast stream is stored somewhere in your files under a name like 0 (e.g.: 1, when recording from [Radio CAFF](http://radiocaff.com.ar/) (or more precisely from the underlying [WinAmp stream](http://panel7.serverhostingcenter.com/tunein.php/radiocaff/playlist.pls)). Under my german Windows XP it was stored under "Eigene Dateien/Eigene Music" so I guess that you find it in an english Windows under "My Documents/My Music/", I don't know where it will be stored under Linux or any other OS (updates are welcome).

You can automagically cut the stream into tracks by relaying the stream through [Streamripper](http://streamripper.sourceforge.net), i.e. by directing StreamRipper to the ShoutCast stream and directing VLC to the relaying port of StreamRipper (default http://localhost:8000).

#### Audio Device

This option helps you to listen to audio files in two modes: stereo and mono.

1.  To listen to an audio track in either the Stereo or Mono mode, select *Open File or Open Disc* from the *Media* menu. The Open dialog box is displayed.
2.  Select an audio file and click on the  *Play* button. The selected track is played.
3.  Select *Mono* in *Audio Device* from the *Audio* menu if you want to listen to the audio track in the Mono mode.

Mono refers to monaural sound that uses a single channel for sound reproduction.

1.  Select *Stereo* in *Audio Device* from the *Audio* menu if you want to listen to the audio track in the Stereo mode.

Stereo refers to sound that uses two channels for sound reproduction or stereophonic sound.

#### Audio Channels

In audio, a channel refers to a stream of audio that is to be played by one speaker. For example, stereo audio, consists of two channels. This option is useful for codecs that don’t have support for more than 2 channels.

Select a channel type in *Audio Channels* from the *Audio* menu. VLC media player provides four audio channels and they are:

1.  *Stereo* – Refers to the reproduction of the sound in two or more independent audio channels using more than one speaker. If you use this option, you would feel as though the sound is played from all the directions. You can observe this in a regular home theatre with 5.1 or 6.1 speakers.
2.  *Left* – You can observe this in a regular audio player with 2.1 speakers. If you select the **Left** option, the music is played only in the left speaker. The speaker on your right is automatically switched OFF.
3.  *Right* - If you select the **Right** option, the music is played only in the speaker on your right side. The speaker on your left is automatically switched OFF.
4.  *Reverse Stereo* – There are several applications that are used to reverse the stereo whereas VLC has an in-built feature to reverse the stereo. This option is useful if you want the audio to play in tandem with the video. You can use the **Reverse Stereo** option if you want to deliberately change the audio output.

Imagine that you are watching a video. In the video, a person walks on the left side but the sound is produced on the right speaker. You can correct this by selecting the *Reverse Stereo* option in VLC. Select the *Reverse Stereo* option and play the same scene in the video and observe the difference.

You can observe this with 2.1, 5.1, 6.1 and 8.1 speakers.

#### Visualize Audio

Visualizations display splashes of colour and geometric shapes and generate animated imagery based on a piece of music.

The different visual effects available are *Spectrometer, Scope, Spectrum, VU Meter and Goom*. This menu item can also be used to disable a visualization.

1.  Select an option under the *Visualizations* option from the *Audio* menu to view the effects. The selected visualization is then played.
2.  To disable visualizations, select *Disable* under *Visualizations* from the *Audio* menu. The visualization is then disabled.

Spectrum visualization on VLC:

#### Maximum VLC Volume

To change the maximum volume in % that VLC should use, go to **Tools** → **Preferences** (select **All** at bottom left corner) → **Interface** → **Main interfaces** → **Qt** → **Maximum volume displayed**.

Save it and restart VLC.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Subtitles {#subtitles}

VLC supports many kinds of subtitles.

#### Media with included subtitles

Many types of media can have embedded subtitles. VLC can read subtitles for the following media formats:

-   DVD
-   SVCD
-   OGM files
-   Matroska (MKV) files

Subtitles are enabled by default in VLC media player. To disable them, go to the *Video* menu, and to *Subtitles track*. All available subtitles tracks will be listed. Select "Disable" to turn off the subtitles. Depending on the media, a description (language, for example) might be available for the track.

To disable subtitles by default, select "Preferences", then "Show All". Select "Input/Codecs". On the "Subtitle Track ID" selection window, change the value to "-1". (NOTE: Changing the value in the "Subtitle Track" menu will not disable the subtitle file.) In the case of multiple subtitle tracks, a value of "0" will enable subtitle track 1, a value of "1" will enable subtitle track 2, and so on.

VLC under Linux:

VLC under OSX:

VLC under Windows:

DVD and SVCD subtitles are merely images, so you won't be able to change anything for them. OGM and Matroska subtitles are rendered text, so you will be able to change several options.

Text rendering options can be changed in the *Preferences* in the *Tools* tab. To adjust the font preference check the *All* bullet in the *Show Settings* box, and then click *Subtitles/OSD*. You can then set the font and its size under *Text Renderer*. For the font, you have to select a font file. In Windows, they can be found in *C:\\Windows\\Fonts*. Under MacOS X, they are in */System/Library/Fonts*. Sizes can be set either relatively or as a number of pixels.

Subtitle text rendering preferences under Windows, VLC 1.1.5

You need to restart your stream for the font modifications to take effect.

#### Subtitles files

While modern file formats like Matroska or OGM can handle subtitles directly, older formats like AVI can't. Therefore, a number of subtitles files formats have been created. You need two files: the video file and the subtitles file that only contains the text of the subtitles and timestamps.

VLC can handle these types of subtitles files:

-   MicroDVD
-   SubRIP
-   SubViewer
-   SSA
-   Sami
-   Vobsub (this one is quite special: it is not made from text but from images, which means that you can't change the fonts)

To open a subtitles file, use the Advanced Open dialog box (Menu File, Open file). Select your file by clicking on the *Browse* button. Then, check the *Subtitle options* checkbox and click on the "Settings" button.

You can then select the subtitles file by clicking the *Browse* button. You can also set a few options like character encoding, alignment and size.

An alternative is loading subtitles from the *Subtitles Track* menu item under the *Video* tab.

Note: For Vobsub subtitles, you need to select the **.idx** file, not the **.sub** file. Encoding, alignment and size won't have any effect for Vobsub subtitles.

Font can be changed as explained in the previous section.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Video {#video}

You can play video files, video clips and other video media using the VLC media player. You can resize, change aspect ratio, crop videos, load subtitles, deinterlace, save snapshots, and convert videos to DirectX wallpaper.

Video tracks of the *.asf, .avi, .divx, .dv, .mxf, .ogg, .gm, .ps, .ts, .vob,* and *.wmv* formats are supported.

#### Playing a Video Track

There are two main ways to open and play a video track:

1.  Select *Open File* from the *Media* menu.

     2. Select a video track and double-click it or click the *Open* button.

The selected track will play.

#### Loading Subtitle Tracks

A subtitle is a textual version of a movie’s dialogue. Subtitles are helpful if you are viewing a movie that contains foreign language(s). You can load subtitles for video tracks. Subtitles of the formats *.cdg, .idx, .srt, .sub, .utf, .ass, .ssa, .aqt, .jss, .psb, .rt* and *smi* are supported.

VLC can read subtitles for the media formats such as *DVD*, *SVCD*, *OGM* files, and *Matroska (MKV)* files.

To enable the subtitle for a track:

1.  Select *Open File* under the *Subtitle* menu item from the *Video* menu. The *Open Subtitles File* dialog box is displayed.  
2.  Locate the file which contains the subtitle and click on *Open*. The subtitles are displayed.

For more details, see [Documentation:Subtitles](#subtitles).

#### Full Screen

This option is useful if you want to watch the video in the full screen mode.

1.  Select *Full Screen* from the *Video* menu. The video will then occupy the entire screen.
2.  To return to the original mode, press *Esc* on the keyboard or right-click the mouse and select the *Leave Full Screen* option. The video will then return to its original mode.

Note: When you switch to full screen, the controls may appear for a short period of time. To restore the controls after they disappear, move the mouse or press any key on the keyboard.

#### Always on Top

This option is useful if you want the VLC media player to remain on the top of the screen always when other applications or files are open.

1.  To make the VLC media player appear on top of the screen, select *Always on Top* from the *Video* menu. 
2.  If you do not want VLC to appear on the top of the screen, select the *Always on Top* option from the *Video* menu and manually minimise the VLC application.

#### DirectX Wallpaper

This option is useful if you want to display the video which is being played as your desktop wallpaper.

To view the current video file as wallpaper

1.  Select *Advanced File Open* from the *Media* menu. The *Open Media* dialog box is displayed. 
2.  Select a file and click  *Play*.
3.  Select *DirectX Wallpaper* from the *Video* menu.

The wallpaper mode will then display the video as the desktop background.

Note: that this feature works only if you deactivate the overlay under Windows XP.

#### Snapshot

This option is useful if you want to capture a portion of the video as an image.

1.  Select *Advanced File Open* from the *Media* menu. The Open dialog box is displayed.
2.  Select a file and click  *Play*.
3.  To capture an image from the video, select *Snapshot* from the *Video* menu.

The image is captured in the *.png* picture format and is saved in the *C:\\My Pictures* folder by default (*C:\\Users\\**Username**\\Pictures*).

#### Zoom

You can enlarge videos in different sizes. This option is useful if you want to change the size of a video track which is being played. The supported sizes are *1:4 Quarter, 1:2 Half, 1:1 Original (default)* and *2:1 Double*.

To view a video in a particular dimension, select a dimension from *Zoom* in the *Video* menu. The track is then resized based on the selected zoom ratio.

#### Aspect Ratio

Aspect ratio refers to the width of a picture in relation to its height. For example, the ratio 4:3 means four units wide to three units high. VLC provides a list of aspect ratio values which are *Default, 1:1, 4:3, 16:9, 16:10, 2.21:1, 2.35:1, 2.39:1* and *5:4*.

To select an aspect ratio, select a value from *Aspect Ratio* in the *Video* menu. The video is then adjusted based on the selected ratio.

#### Crop

This option is helpful if you want to capture a small portion of a video as an image. This also helps crop the black bars of the top and bottom of a video.

The cropping values that are supported are *Default, 16:10, 16:9, 1.85:1, 2.21:1, 2.35:1, 2.39:1, 5:3, 4:3, 5:4,* and *1:1*.

To crop a video that is played, select a value from *Crop* in the *Video* menu. The video is then cropped based on the selected value.

#### Deinterlace

Deinterlace refers to a process where interlaced video signals are converted into non-interlaced signals. VLC provides the *Discard, Blend, Mean, Bob, Linear, X, Yadif and Yadif (2x)* deinterlacement modes.

1.  Select *Deinterlace* from the *Video* menu and choose the appropriate setting.
2.  To change the deinterlacement mode select 'Deinterlace mode' is the *Video Menu*
3.  Select a mode and observe the change in the video being played.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Video and Audio Filters {#video-and-audio-filters}

This page is outdated and information might be incorrect.

VLC includes a system of *filters* that allow you to modify the audio and video.

#### Deinterlacement and Post Processing

VLC is able to deinterlace a video stream using different deinterlacement methods. Deinterlacement can be enabled in the *Video* menu, *Deinterlacement* menu item. The *Blend* methods gives the best results in most cases. The *discard* method is a less resource consuming alternative, although its results may be slightly compromised.

On some particular streams (MPEG 4, DivX, Xvid, Sorenson, etc.), some additional image filtering can be applied to the video before display, improving its quality in some cases. This can be enabled by using the *Post processing* menu item in *Video*. Different levels of post processing can be chosen here. A higher level means more filtering.

#### Video filters

##### Summary

VLC features several filters able to change the video (distortion, brightness adjustment, motion blurring, etc.).

In Windows and Linux, the user must go to the *Effects and Filters* in the *Tools* menu item. A dialogue box entitled "Adjustments and Effects" will appear.

In macOS you can enable these filters through the *Extended Controls panel*. Click on the triangle next to *Video filters* to select your filters or expand the *Adjust Image* section to change the contrast, hue, etc.

iOS:

Example of combined effects on a video:

##### Rotate

You can easily rotate a video. Open the *Effects and Filters* dialog, in the *Tools menu*

Select the *Video Effects* tab, then the *Geometry* one.

Check the *Transform* checkbox to use rotation presets (90°, 180°, 270°) or check the *Rotate* checkbox to manually select the angle you wish to apply.

#### Audio filters

##### Equalizer

[Wikipedia](http://en.wikipedia.org/wiki/Main_Page) has information on this entry:

***[Equalization (audio)](http://en.wikipedia.org/wiki/Equalization_(audio) "wikipedia:Equalization (audio)")***

VLC features a 10-band graphical equalizer, a device used to alter the relative frequencies of audio (e.g. for a bass boost). You can display it by activating the advanced GUI on wxWidgets or by clicking the *Equalizer* button on the macOS interface. The following image is the interface of the audio equalizer in the Windows and GNU/Linux interface.

The equalizer in the macOS interface

Presets are available in all of these dialog boxes.

##### Other audio filters

At the moment, VLC features two other audio filters: a volume normalizer and a filter providing sound spatialization with a headphone. They can be enabled in the *Effects and Filters* menu item in the *Tools* tab of the Windows and GNU/Linux interface and in the Audio section of the Extended Controls panel of the macOS interface.

For better control, you need to go to the preferences. To select the filters to be enabled, go to *Audio*, then to *Filters*. In the "audio filters" box, enter the names of the filters to enable, separated by commas. Valid names are "equalizer", "normvol" and "headphone".

If you want to tune the behavior of these filters, go to *Audio, Filters, \[your filter\]*. The equalizer and headphone filters can be tuned.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

## Getting Started {#getting-started}

### Install VLC {#install-vlc}

There are VLC binaries available for the many OSes, but not for all supported ones. If there are no binaries for your OS or if you want to change the default settings, you can compile VLC from source.

####  Windows

#### 95, 98, ME

You can install VLC on Windows 95, 98, or ME operating systems by using [KernelEx](http://kernelex.sourceforge.net/wiki/Main_Page).

#### 2000, XP, Vista, 7, 8, 10

##### Recommended

The normal and recommended way to install VLC on a Windows operating system is via the installer package.

**Step 0: Download and launch the installer**

Download the installer package from the [VLC download page for Windows](https://www.videolan.org/vlc/download-windows.html). After you download the installer package, double click on the file to begin the install process. If you're using Windows Vista, 7, 8 or 10 and have UAC (User Account Control) enabled, the operating system may prompt you to grant VLC administrator permissions. Click **Yes** to continue the installation process.

**Step 1: Select an installer language**

Before you can continue, you must select the language that you want the installer to use to display information to you. After you select a language, click **OK**.

**Step 2: Review the Welcome screen**

The VLC installer recommends that you close all other applications before continuing the installation process. When you're ready to proceed with the installation process, click **Next**.

**Step 3: Read License agreement**

Read the Terms of Service. Once you're done reading, click **Next**.

**Step 4: Select components**

Use this menu to customize your install. Choose all of the components you wish to install and whether you want VLC to be your default media player or not. Once you are done, click **Next**.

**Step 5: Pick a location**

Click **Browse...** to choose the destination installation folder. After you've identified the desired folder, click **Install**.

**Step 6: Now installing**

Wait as VLC is installed on your machine. It shouldn't take too long. Then click "Show details" to see more information about the progress of the installation.

**Step 7: Installation complete**

Once installation is complete, you may choose to run VLC or read VLC's release notes. Click **Finish** to complete the installation process and close the installer.

##### Alternative

If you want to perform an unattended (or silent) installation of VLC, you can do so via a command-line interface. Type in "*filename*" /L="*languagecode*" /S. For example, the English installation would look something like **vlc-2.0.1-win32.exe /L=1033 /S**.

**PowerShell**

Installing VLC using PowerShell is as easy.

**Command Prompt**

You can also install VLC using the command prompt.

####  macOS

1.  Download the macOS package from the [VLC macOS download page](https://www.videolan.org/vlc/download-macosx.html).
2.  Double-click on the icon of the package: an icon will appear on your Desktop, right beside your drives.
3.  Open it and drag the VLC application from the resulting window to the place where you want to install it (it should be **/Applications**).

Note: You may need to delete older versions of VLC on your computer before you can successfully install the latest version.

#### Linux

#####  Debian

Download page: 0

**A standard install without libdvdcss:**

    # apt-get update
    # apt-get install vlc

Or search for 0 with the graphical package manager you like best. It should be in the main Debian repository in the section *Video software*. Additional plugins are available and most require manual selection, e.g. 1, 2 and 3.

**For a standard install with libdvdcss:**

A simple install of the libdvdcss package can be found here: 0, but for future bug fixes add the following lines to your **/etc/apt/sources.list**:

     deb 0 stable main
     deb-src 0 stable main

Then:

    # apt-get update
    # apt-get install vlc libdvdcss2

This will allow you to decrypt DVDs.

######  Ubuntu

Links: [Download page](https://www.videolan.org/vlc/download-ubuntu.html) • Launchpad ([Source](https://launchpad.net/ubuntu/+source/vlc/) • [Bugs sorted by most users](https://bugs.launchpad.net/ubuntu/+source/vlc/+bugs?field.searchtext=&orderby=-users_affected_count&search=Search&field.status%3Alist=NEW&field.status%3Alist=CONFIRMED&field.status%3Alist=TRIAGED&field.status%3Alist=INPROGRESS&field.status%3Alist=FIXCOMMITTED&field.status%3Alist=INCOMPLETE_WITH_RESPONSE&field.status%3Alist=INCOMPLETE_WITHOUT_RESPONSE&assignee_option=any&field.assignee=&field.bug_reporter=&field.bug_commenter=&field.subscriber=&field.tag=&field.tags_combinator=ANY&field.status_upstream-empty-marker=1&field.upstream_target=&field.has_cve.used=&field.omit_dupes.used=&field.omit_dupes=on&field.affects_me.used=&field.has_patch.used=&field.has_branches.used=&field.has_branches=on&field.has_no_branches.used=&field.has_no_branches=on&field.has_blueprints.used=&field.has_blueprints=on&field.has_no_blueprints.used=&field.has_no_blueprints=on) • [Questions](https://answers.launchpad.net/ubuntu/+source/vlc))

Launch the Ubuntu Software Center and go to **All Software → Sound & Video** then in search VLC Player. After it will come click on it and it will automatically install

You need to check that a universe mirror is listed in your **/etc/apt/sources.list** file.

    $ sudo apt-get update
    $ sudo apt-get install vlc vlc-plugin-pulse mozilla-plugin-vlc

As given by 0:

    $ sudo apt install libdvd-pkg && sudo dpkg-reconfigure libdvd-pkg

will install a packaged version of libdvdcss without the need for third-party repos.

##### Red Hat

-
-

Adapted (annotated) from 0:

Red Hat/CentOS/Scientific Linux have almost the same setups (they're all derived from Red Hat). Red Hat and derivatives have [different instructions](https://fedoraproject.org/wiki/EPEL#Quickstart) if EPEL (Extra Packages for Enterprise Linux) is not set up. Red Hat Network (RHN) users should verify that they have enabled the *optionals* and *extras* channels for RHN subscriptions.

If you want to have DVD playback ability, you will need to install the libdvdcss package too ([source](https://www.videolan.org/vlc/download-redhat.html)).

For the latest version (up to the now-current version 3.0.6) use [RPM Fusion](https://rpmfusion.org/RPM%20Fusion), otherwise VLC branches 2.0.x and 2.2.x are available: Red Hat/CentOS/Scientific Linux 7: (vlc-2.2.x – branch available for x86_64 architectures)

    $> su -
        #> yum install 0
        #> yum install 0
        #> yum install vlc
        #> yum install vlc-core             # (for minimal headless/server install)
        #> yum install python-vlc npapi-vlc # (optionals)

Red Hat/CentOS/Scientific Linux 6: (vlc-2.0.x branch – available for i686 and x86_64 architectures)

    $> su -
        #> yum install 0
        #> yum install 0
        #> yum install vlc
        #> yum install vlc-core             # (for minimal headless/server install)
        #> yum install python-vlc npapi-vlc # (optionals)

##### SUSE

Download page: 0

#### FreeBSD

Download page: 0

Install vlc from the packages collection:

    # pkg install vlc

#### Compile the sources by yourself

For more detailed information on compiling VLC, please see Compile VLC.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Installing VLC {#installing-vlc}

There are VLC binaries available for the many OSes, but not for all supported ones. If there are no binaries for your OS or if you want to change the default settings, you can compile VLC from source.

####  Windows

#### 95, 98, ME

You can install VLC on Windows 95, 98, or ME operating systems by using [KernelEx](http://kernelex.sourceforge.net/wiki/Main_Page).

#### 2000, XP, Vista, 7, 8, 10

##### Recommended

The normal and recommended way to install VLC on a Windows operating system is via the installer package.

**Step 0: Download and launch the installer**

Download the installer package from the [VLC download page for Windows](https://www.videolan.org/vlc/download-windows.html). After you download the installer package, double click on the file to begin the install process. If you're using Windows Vista, 7, 8 or 10 and have UAC (User Account Control) enabled, the operating system may prompt you to grant VLC administrator permissions. Click **Yes** to continue the installation process.

**Step 1: Select an installer language**

Before you can continue, you must select the language that you want the installer to use to display information to you. After you select a language, click **OK**.

**Step 2: Review the Welcome screen**

The VLC installer recommends that you close all other applications before continuing the installation process. When you're ready to proceed with the installation process, click **Next**.

**Step 3: Read License agreement**

Read the Terms of Service. Once you're done reading, click **Next**.

**Step 4: Select components**

Use this menu to customize your install. Choose all of the components you wish to install and whether you want VLC to be your default media player or not. Once you are done, click **Next**.

**Step 5: Pick a location**

Click **Browse...** to choose the destination installation folder. After you've identified the desired folder, click **Install**.

**Step 6: Now installing**

Wait as VLC is installed on your machine. It shouldn't take too long. Then click "Show details" to see more information about the progress of the installation.

**Step 7: Installation complete**

Once installation is complete, you may choose to run VLC or read VLC's release notes. Click **Finish** to complete the installation process and close the installer.

##### Alternative

If you want to perform an unattended (or silent) installation of VLC, you can do so via a command-line interface. Type in "*filename*" /L="*languagecode*" /S. For example, the English installation would look something like **vlc-2.0.1-win32.exe /L=1033 /S**.

**PowerShell**

Installing VLC using PowerShell is as easy.

**Command Prompt**

You can also install VLC using the command prompt.

####  macOS

1.  Download the macOS package from the [VLC macOS download page](https://www.videolan.org/vlc/download-macosx.html).
2.  Double-click on the icon of the package: an icon will appear on your Desktop, right beside your drives.
3.  Open it and drag the VLC application from the resulting window to the place where you want to install it (it should be **/Applications**).

Note: You may need to delete older versions of VLC on your computer before you can successfully install the latest version.

#### Linux

#####  Debian

Download page: 0

**A standard install without libdvdcss:**

    # apt-get update
    # apt-get install vlc

Or search for 0 with the graphical package manager you like best. It should be in the main Debian repository in the section *Video software*. Additional plugins are available and most require manual selection, e.g. 1, 2 and 3.

**For a standard install with libdvdcss:**

A simple install of the libdvdcss package can be found here: 0, but for future bug fixes add the following lines to your **/etc/apt/sources.list**:

     deb 0 stable main
     deb-src 0 stable main

Then:

    # apt-get update
    # apt-get install vlc libdvdcss2

This will allow you to decrypt DVDs.

######  Ubuntu

Links: [Download page](https://www.videolan.org/vlc/download-ubuntu.html) • Launchpad ([Source](https://launchpad.net/ubuntu/+source/vlc/) • [Bugs sorted by most users](https://bugs.launchpad.net/ubuntu/+source/vlc/+bugs?field.searchtext=&orderby=-users_affected_count&search=Search&field.status%3Alist=NEW&field.status%3Alist=CONFIRMED&field.status%3Alist=TRIAGED&field.status%3Alist=INPROGRESS&field.status%3Alist=FIXCOMMITTED&field.status%3Alist=INCOMPLETE_WITH_RESPONSE&field.status%3Alist=INCOMPLETE_WITHOUT_RESPONSE&assignee_option=any&field.assignee=&field.bug_reporter=&field.bug_commenter=&field.subscriber=&field.tag=&field.tags_combinator=ANY&field.status_upstream-empty-marker=1&field.upstream_target=&field.has_cve.used=&field.omit_dupes.used=&field.omit_dupes=on&field.affects_me.used=&field.has_patch.used=&field.has_branches.used=&field.has_branches=on&field.has_no_branches.used=&field.has_no_branches=on&field.has_blueprints.used=&field.has_blueprints=on&field.has_no_blueprints.used=&field.has_no_blueprints=on) • [Questions](https://answers.launchpad.net/ubuntu/+source/vlc))

Launch the Ubuntu Software Center and go to **All Software → Sound & Video** then in search VLC Player. After it will come click on it and it will automatically install

You need to check that a universe mirror is listed in your **/etc/apt/sources.list** file.

    $ sudo apt-get update
    $ sudo apt-get install vlc vlc-plugin-pulse mozilla-plugin-vlc

As given by 0:

    $ sudo apt install libdvd-pkg && sudo dpkg-reconfigure libdvd-pkg

will install a packaged version of libdvdcss without the need for third-party repos.

##### Red Hat

-
-

Adapted (annotated) from 0:

Red Hat/CentOS/Scientific Linux have almost the same setups (they're all derived from Red Hat). Red Hat and derivatives have [different instructions](https://fedoraproject.org/wiki/EPEL#Quickstart) if EPEL (Extra Packages for Enterprise Linux) is not set up. Red Hat Network (RHN) users should verify that they have enabled the *optionals* and *extras* channels for RHN subscriptions.

If you want to have DVD playback ability, you will need to install the libdvdcss package too ([source](https://www.videolan.org/vlc/download-redhat.html)).

For the latest version (up to the now-current version 3.0.6) use [RPM Fusion](https://rpmfusion.org/RPM%20Fusion), otherwise VLC branches 2.0.x and 2.2.x are available: Red Hat/CentOS/Scientific Linux 7: (vlc-2.2.x – branch available for x86_64 architectures)

    $> su -
        #> yum install 0
        #> yum install 0
        #> yum install vlc
        #> yum install vlc-core             # (for minimal headless/server install)
        #> yum install python-vlc npapi-vlc # (optionals)

Red Hat/CentOS/Scientific Linux 6: (vlc-2.0.x branch – available for i686 and x86_64 architectures)

    $> su -
        #> yum install 0
        #> yum install 0
        #> yum install vlc
        #> yum install vlc-core             # (for minimal headless/server install)
        #> yum install python-vlc npapi-vlc # (optionals)

##### SUSE

Download page: 0

#### FreeBSD

Download page: 0

Install vlc from the packages collection:

    # pkg install vlc

#### Compile the sources by yourself

For more detailed information on compiling VLC, please see Compile VLC.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Quick Start Guide {#quick-start-guide}

**Languages:English** • Deutsch

Read on for a quick overview of VLC's features and capabilities.

#### Starting VLC

##### Windows

-   In Windows XP: Click **Start** -\> **Programs** -\> **VideoLAN** -\> **VLC media player**.
-   In Windows 7: Click **Start** -\> **All Programs** -\> **VideoLAN** -\> **VLC media player**.

VLC is shown on the screen and a small icon  is shown in the system tray.

##### macOS

Start VLC from the applications menu or the system dock.

VLC is shown on the screen and a small icon  is shown in the dock.

##### Linux

Start VLC from the command line with **vlc** or start it from your desktop environment's application launcher.

#### Interface

##### The main interface

-   VLC media player on Windows and Linux: **VLC media player on macOS**

##### More interface informations

Go to [Documentation:Interface](#interface)

#### Play a media

##### Play a single media file

Find a media file you want to play with your favourite File Explorer (Windows Explorer, Finder, Konqueror...) and double-click on it.

You can also drag and drop the file onto VLC.

##### Play a whole media folder

Start VLC, open the *Media* menu, and select the *Open Folder...* menu item. An *Open Folder* dialog box will appear. Select the folder you want to open and select *Open*.

##### Play a CD/DVD/VCD

Insert your disk and your OS should ask you what you want to do. Select *Play with VLC* and select the OK button.

##### More open options

Go to [Documentation:Open Media](#open-media)

#### Preferences

##### Where are the VLC preferences?

To open the *Preferences* panel, open the Tools menu , and select the *Preferences* menu item.

Here is the Simple Preferences panel where you can modify the essential settings of VLC.

##### How to reset the VLC preferences?

Go to VSG:ResetPrefs

#### Playlist view

##### Overview

This view allows you to easily browse different sources of media. To access the Playlist View, click on the *Playlist* button in the main interface.

-   **1:**: The current Playlist you are listening and your Media Library

-   **2:**: The OS default media folders

-   **3:**: Your local optic drive (CD, DVD...)

-   **4:**: Your local network sources

-   **5:**: Internet sources (Podcasts, Shoutcast radios...)

-   **6:**: The media listing you are listening or browsing


##### More Playlist options

Go to [Documentation:Playlist](#playlist)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Uninstalling VLC {#uninstalling-vlc}

#### Windows

You can uninstall VLC from *Add/Remove Programs* (*Programs and Features* in Windows 7) located in the *Control Panel*. Search for VLC media player and right click, then select "Uninstall/Change". Follow the prompts to finish the uninstallation.

Alternatively, you can browse to VLC's installation directory (for a typical install, go to your *C: Drive* and look for Program Files (if 64-bit, Program Files (x86) )→VideoLAN→VLC and double-click on the *uninstall* link and follow the prompts to uninstall.

#### macOS

Drag the VLC application to your trash can. You can also remove the configuration file and the cache files in **\~/Library/Preferences/VLC/**. There is an AppleScript on the disk-image which lets you do this automatically.

If that did not work, you can double-click on the *Applications* icon. This will bring up a list of all applications on your Mac. Scroll through the list of Applications, then press and hold the *Ctrl* button to bring up a table of options and actions. Click on "move to trash".

Finally, if the previous processes failed, you can try downloading a third-party uninstaller program to uninstall it, such as [AppCleaner](http://www.macupdate.com/app/mac/25276/appcleaner).

#### Linux

##### Debian

Remove the packages that you installed:

    # apt-get remove --purge vlc libdvdcss2

###### Ubuntu

Remove *VLC Media Player* by entering this command in the Terminal.

    $  sudo apt-get remove vlc

Or you can also search *VLC* in the *Ubuntu Software Center* and click on *Remove* to uninstall it.

##### Red Hat and SuSE

Uninstall the RPM packages that you installed:

    # rpm -e vlc-version vlc-mad-version wxvlc-version libdvdcss2-version libdvdpsi1-version

#### Compiled the sources by yourself

Go to the directory containing VLC sources and execute

    # make uninstall

You can then remove the VLC sources.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### User Guide {#user-guide}

This is the user guide for the VLC media player.

#### VLC User Guide

-   Quick start guide: How to start with VLC.
-   [Installation](#installing-vlc): Installation instructions for several systems.
-   [History](#history): Overview and history of the VideoLAN project.

#### Usage

-   [Interface](#interface): The main interface of VLC media player.
    -   OSX Interface
    -   Windows/Linux Interface
-   [Open Media](#open-media): Open every media you want, the way you want.
-   [Audio](#audio): Visualization, selection of devices...
-   [Video](#video): Cropping, snapshots and screenshots...
-   [Playback](#playback): Navigation through media files (e.g. chapters, bookmarks).
-   [Playlist](#playlist): Creating and managing playlists.
-   [Subtitles](#subtitles): Selection of subtitles
-   [Video and Audio Filters](#video-and-audio-filters): Usage of VLC's filters (equalizer, video filters)
-   [Snapshots](#snapshots): How to create snapshots and screenshots.
-   [Hotkeys](#hotkeys): Configuration of VLC's hotkeys
-   Uninstallation: Uninstallation instructions.
-   Troubleshooting: The VLC Support Guide, an informal, step-by-step guide for troubleshooting most common issues with VLC.

#### Advanced Usage

-   [Using VLC inside a webpage](#webplugin): How to create webpages that use the VLC Web plugin.
-   [Command line](#command-line): Main command line instructions.
-   [Alternative Interfaces](#alternative-interfaces) : HTTP interface and other control interface.
-   [Misc](#misc) : Miscellaneous other things.

#### Appendix

-   Building Pages for the HTTP Interface
-   Format String
-   Building Lua Playlist Scripts
-   VLC Use 0.8. (Versions older than 0.9).

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### VLC for dummies {#vlc-for-dummies}

**Languages:English** • Deutsch • Nederlands

Thank you for visiting this page to find out what VLC media player is all about. **Please Note**: this page is *Under Construction* and might change in the near future.

VLC media player (or VLC for short) is a FREE and Open Source Software Media Player. Features that make VLC the preferred media player for a lot of people are its excellent support for various Audio and Video codecs, the fact that it's cross platform and the open way of development.

This page describes the basic use of VLC. See [VLC Play HowTo](#play-howto) for a user guide and [VLC Streaming HowTo](#streaming-howto) for advanced streaming features.

#### Prerequisites

To use VLC you need:

-   A computer with [Internet](http://en.wikipedia.org/wiki/Internet) access and an internet (web) browser (to [download VLC](http://www.videolan.org/vlc/#download)).
    [What is a web browser?](http://www.youtube.com/watch?v=BrXPcaRlBqo)
-   Media (audio or video) [files](http://en.wikipedia.org/wiki/Computer_file) or [Disc](http://en.wikipedia.org/wiki/Optical_disc) ([optical drive](http://en.wikipedia.org/wiki/Optical_disc_drive) required to play discs).
-   Audio output hardware (speakers, headphones) for audio playback.
-   Knowledge about working with computer files and folders.
    [Working with files and folders](http://windows.microsoft.com/en-US/windows7/Working-with-files-and-folders) ([MSFT Windows](http://windows.microsoft.com))
    [How to open documents and folders](http://support.apple.com/kb/PH3725) ([macOS](http://www.apple.com/macos/))
    [Files, folders & search](https://help.ubuntu.com/11.10/ubuntu-help/files.html) ([Ubuntu](http://ubuntu.com))

#### The main interface

##### Interface overview

The following picture shows the names of the main controls in the VLC interface:

**Note:** This picture corresponds to the Windows XP version. In other systems, VLC might look slightly different.

##### Menu bar

The menu bar at the top contains commands that control VLC.

##### Track slider

The track slider is on top of the control buttons. It shows the progress of playing of the media file. You can drag the track slider left to rewind or right to forward the track being played.

Two timers at the left and right ends of the track slider show the current playing position (left) and the total time (right) of the current track.

**Note:** When a media file is streamed (live), the position indicator of the track slider does not move because the total duration of the streaming is not known until it finishes.

##### Control buttons

The buttons below the slider control the playback.

From left to right they are:

-   Play/Pause.
-   Previous media in the playlist.
-   Stop playback.
-   Next media in the playlist.
-   Toggle fullscreen (video only).
-   Show extended settings: Audio effects, Video effects and Synchronization.
-   Show playlist.
-   Repeat: toggles among loop all, loop one, no loop (default).
-   Random: Plays the files in the current playlist in a random order.

##### Volume control

The volume control is located in the bottom right corner of the window. The small speaker icon is a button that mutes () or un-mutes () the sound. The triangle to the right is a slider that shows the current playback volume. Clicking this slider modifies the volume. The playback volume is also displayed as a percentage number on top of this slider.

#### Windows notification area (system tray) icon

When you start VLC media player, the application appears on the screen and a small icon  appears in the notification area (system tray). Clicking once this icon will hide VLC, and clicking it again will show it again. Hiding VLC does not close it, it continues to run in the background. Right clicking this icon brings up a menu with the following controls:

-   *Hide/Show VLC media player*.
-   *Play/Pause/Stop playback*.
-   *Switch to Previous/Next track*.
-   *Speed control*.
-   *Increase/Decrease volume*.
-   *Mute*.
-   *Open media*.
-   *Quit*.

#### Tutorials

-   Installing VLC
    -   Windows
    -   macOS
    -   Linux
-   Basics of VLC
    -   Windows
        -   starting VLC
            Double click the VLC icon  on the desktop or from the start menu: select *Programs*, select *VideoLAN* and select *VLC media player*.
    -   Playing media files stored in the computer
        -   Queuing files
            You can queue files by selecting multiple files at a time.
    -   Playing media from your optical reader (CD, DVD, Blu-Ray)
    -   Closing VLC

#### See also

-   VLC HowTo
-   [Documentation:Play HowTo](#play-howto)

&nbsp;

-   Common Problems
-   Hotkeys
    -   How to set global hotkeys

#### External Links

-   [F.A.Q.](http://www.videolan.org/support/faq.html)
-   [Windows Help and How-to](http://windows.microsoft.com/en-US/windows/help)
-   [What is a web browser?](http://www.youtube.com/watch?v=BrXPcaRlBqo)
-   [Video tutorials (selected by Jan)](http://www.youtube.com/playlist?list=PLCCCDDDC322A8AA82)

**Languages:English** • Deutsch • Nederlands

## Modules and Options {#modules-and-options}

### a52 {#modules-a52}

See also: Documentation:Modules/es

Module: a52

**Type**: Access demux

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: ATSC A/52 (AC-3) audio decoder

**Shortcut(s)**: (none)

There is a comment in the code:

    /*
     * NOTA BENE: this module requires the linking against a library which is
     * known to require licensing under the GNU General Public License version 2
     * (or later). Therefore, the result of compiling this module will normally
     * be subject to the terms of that later license.
     */

#### Options

-   **a52-dynrng \** : Dynamic range compression makes the loud sounds softer, and the soft sounds louder, so you can more easily listen to the stream in a noisy environment without disturbing anyone. If you disable the dynamic range compression the playback will be more adapted to a movie theater or a listening room *default value: enabled*

#### Source code

-   [modules/codec/a52.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/a52.c)
-   [modules/packetizer/a52.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/packetizer/a52.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### adjust {#modules-adjust}

Module: adjust

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Image properties filter

**Shortcut(s)**: -

**Note:** Before version 0.9.0, this used to be a vout filter.

#### Options

-   **contrast \** : Contrast *default value: 1.0*
-   **brightness \** : Brightness *default value: 1.0*
-   **hue \** : Hue *default value: 0*
-   **saturation \** : Saturation *default value: 1.0*
-   **gamma \** : Gamma *default value: 1.0*
-   **brightness-threshold \** : When this mode is enabled, pixels will be shown as black or white. Also may invert the brightness value. The threshold value will be the brightness defined below *default value: disabled*

#### Examples

    % vlc --video-filter "adjust{hue=120,gamma=2.}" somevideo.avi

#### Source code

-   [modules/video_filter/adjust.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/adjust.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### adpcm {#modules-adpcm}

Module: adpcm

**Type**: Audio decoder

**First VLC version**: 0.5.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: ADPCM audio decoder

**Shortcut(s)**: (none)

#### Options

None.

#### Source code

-   [modules/codec/adpcm.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/adpcm.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### alphamask {#modules-alphamask}

**Mosaic framework (How-To)Modules:** mosaic (mosaic-bridge • bridge-in • bridge-out) • alphamask • bluescreen

Module: alphamask

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Change the video's alpha channel

**Shortcut(s)**: 0, 1

This filter can be used in the mosaic framework to set a video's alpha channel (or transparency) based on a PNG image's alpha channel. You can thus blend only parts of the mosaic substream on the background.

#### Options

-   **alphamask-mask \** : PNG file to use as a mask. The alpha channel only will be used to build the mask. This image needs to have the same size as the video it will be used with. *default value: NULL*

#### Example

    % vlc -I telnet --color -vvv --vlm-conf mosaic.vlm --mosaic-keep-picture
      --fake-file ~/images/mire.jpg --fake-width 360 --fake-height 270
      --no-audio --sub-filter mosaic

And the vlm config:

    new channel0 broadcast enabled
    setup channel0 input redefined-nintendo.mpg
    setup channel0 output #duplicate{dst=mosaic-bridge{height=270,width=360,chroma=YUVA,vfilter=alphamask{mask=cone_360x270.png}},select=video}

    new background broadcast enabled
    setup background input fake:
    control background play

    control channel0 play

The files used are available on [people.videolan.org/\~dionoea/mosaic (archived)](https://web.archive.org/web/20121015070412/0) if you want to test. (This will blend the redefined nintendo video in a cone like region)

#### Source code

-   [modules/video_filter/alphamask.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/alphamask.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### alsa {#modules-alsa}

oss and alsa audio capture support were removed from v4l and v4l2 in VLC 1.0.0, but accesses were provided as sub-modules. To emulate old behaviour, use 0 or 1.

In the module options below 0{.variable} and other variables are defined in [include/vlc_es.h](https://git.videolan.org/?p=vlc.git;a=blob;f=include/vlc_es.h). The values are not defined here because of their complexity.

Audio channels in VLC 2.0.1 must be configured manually (bugs) but 0 defaults to stereo.

The access module option 0 has been deprecated since VLC 2.1.0.

HDMI support is planned for VLC 4.0.0 through the 0 option.

#### Options

##### Audio output

Module: alsa

**Type**: Audio output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: ALSA audio output

**Shortcut(s)**: (none)

-   **alsa-audio-device \** : Audio output device (using ALSA syntax) *default value: default*
-   **alsa-audio-channels \** : Channels available for audio output. If the input has more channels than the output, it will be down-mixed. This parameter is ignored when digital pass-through is active *default value: 0{.variable}*

##### Access

Module: alsa

**Type**: Access

**First VLC version**: 1.0.0

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: ALSA audio capture

**Shortcut(s)**: 0

-   **alsa-stereo \** : Stereo *default value: enabled*
-   **alsa-samplerate \ 192000, 176400, 96000, 88200, 48000, 44100, 32000, 22050, 24000, 16000, 11025, 8000, 4000** : Sample rate (Hertz) *default value: 48000*

#### Source code

-   [modules/audio_output/alsa.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/audio_output/alsa.c)
-   [modules/access/alsa.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/alsa.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### anaglyph {#modules-anaglyph}

Module: anaglyph

**Type**: Video filter

**First VLC version**: 2.1.0

**Last VLC version**: -

**Operating system(s)**: Cross platform

**Description**: Renders 3D videos using anaglyph technology

**Shortcut(s)**: -

VLC, since version 2.1.0, is capable of playing 3D side-by-side (SBS) videos using anaglyph technology, viewable with typical paper red-cyan 3D glasses.

To enable this filter, go to **Tools → Effects and Filters → Video Filters → Advanced** and check off **Anaglyph 3D**.

The colour scheme of the glasses can be changed by going to **Tools → Preferences (All) → Video → Filters → Anaglyph** and changing the combo box labelled **Color scheme** on the right. Save and restart VLC for settings to take effect.

#### Options

-   **anaglyph-scheme \** : Define the colour scheme of the glasses. Possible options are:

\- **red-green** - pure red (left) pure green (right)
- **red-blue** - pure red (left) pure blue (right)
- **red-cyan** - pure red (left) pure cyan (right)
- **trioscopic** - pure green (left) pure magenta (right)
- **magenta-cyan** - magenta (left) cyan (right)
*default value: red-cyan*

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### arts {#modules-arts}

Module: arts

**Type**: Audio output

**First VLC version**: -

**Last VLC version**: 0.9.10

**Operating system(s)**: Linux

**Description**: [aRts](http://en.wikipedia.org/wiki/aRts) audio output

**Shortcut(s)**: -

aRts and esd were removed prior to 1.0.0, because both projects were inactive.

Modern Linux users can use the pulse or jack modules instead. There are probably others.

#### Options

None.

#### Source code

-   [modules/audio_output/arts.c](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=blob;f=modules/audio_output/arts.c) (vlc/vlc-0.9.git)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### asf {#modules-asf}

Module: asf

**Type**: Muxer

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: ASF muxer

**Shortcut(s)**: -

Shortcuts to this module are 0 and 1. Support for demuxing was added in 0.5.0. Support for muxing was added sometime prior to 0.8.0, as the changelog says "Improved ASF muxer" for 0.8.0. Support for images/cover art was added in 2.0.0.

#### Options

-   **sout-asf-title \** : Title to put in ASF comments *default value: ""*
-   **sout-asf-author \** : Author to put in ASF comments *default value: ""*
-   **sout-asf-copyright \** : Copyright string to put in ASF comments *default value: ""*
-   **sout-asf-comment \** : Comment to put in ASF comments *default value: ""*
-   **sout-asf-rating \** : "Rating" to put in ASF comments *default value: ""*
-   **sout-asf-packet-size \** : ASF packet size *default value: 4096*
-   **sout-asf-bitrate-override \** : Do not try to guess ASF bitrate. Setting this, allows you to control how Windows Media Player will cache streamed content. Set to audio+video bitrate in bytes *default value: 0*

#### Source code

-   [modules/mux/asf.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/mux/asf.c) (muxer)
-   [modules/demux/asf](https://git.videolan.org/?p=vlc.git;a=tree;f=modules/demux/asf;hb=HEAD)
-   [modules/demux/asf/asf.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/asf/asf.c) (demuxer)
-   [modules/demux/asf/libasf.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/asf/libasf.c) (stream demuxer)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### atmo {#modules-atmo}

Module: atmo

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: 2.2.8

**Operating system(s)**: all

**Description**: outputs colour data via a serial connection to a Atmolight system

**Shortcut(s)**: -

This filter analyzes the running video and outputs color data to drive a Atmolight system.

#### Options

#### Example

#### See also

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### audioqueue {#modules-audioqueue}

Module: audioqueue

**Type**: Audio output

**First VLC version**: 2.0.0

**Last VLC version**: 2.2.8

**Operating system(s)**: macOS

**Description**: AudioQueue (iOS / Mac OS) audio output

**Shortcut(s)**: -

This module was rewritten prior to 2.1.0. It had a single shortcut of 0. It was replaced with auhal (AudioUnit).

#### Options

None.

#### Source code

-   [modules/audio_output/audioqueue.c](https://git.videolan.org/?p=vlc/vlc-2.2.git;a=blob;f=modules/audio_output/audioqueue.c) (vlc/vlc-2.2.git)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### auhal {#modules-auhal}

Module: auhal

**Type**: Audio output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: macOS

**Description**: HAL AudioUnit output

**Shortcut(s)**: -

The option 0 is obsolete since VLC 2.2.0

#### Options

-   **auhal-volume \** : Audio volume *default value: 256*
-   **auhal-audio-device \** : Last audio device *default value: ""*
-   **auhal-warned-devices \** : NULL *default value: ""*

#### Source code

-   [modules/audio_output/auhal.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/audio_output/auhal.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### autodel {#modules-autodel}

Module: autodel

**Type**: Stream output

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Automatically add/delete input streams

**Shortcut(s)**: 0

#### Options

None.

#### Source code

-   [modules/stream_out/autodel.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/stream_out/autodel.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### avcapture {#modules-avcapture}

Module: avcapture

**Type**: Access

**First VLC version**: 2.1.0

**Last VLC version**: -

**Operating system(s)**: macOS

**Description**: AVFoundation video capture module

**Shortcut(s)**: -

The qtcapture module was removed prior to 3.0.0, and users were directed to avcapture.

#### Options

None.

#### Source code

-   [modules/access/avcapture.m](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/avcapture.m)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### avcodec {#modules-avcodec}

See also: Documentation:Modules/avformat

Module: avcodec

**Type**: codec library

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Various audio and video decoders/encoders delivered by the FFmpeg library. This includes (MS)MPEG4, DivX, SVQ1, H261, H263, H264, WMV, WMA, AAC, AMR, DV, MJPEG and other codecs

**Shortcut(s)**: -

libavcodec provided by the FFmpeg project. A full list of supported codecs may be found with [modules/codec/avcodec/fourcc.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/avcodec/fourcc.c)

Options prefixed with *ffmpeg-* or *sout-ffmpeg-* were deprecated in 2.1.0 to reflect the new module name *avcodec*. The only option that seems not to have been replaced later is 0, removed in 3.0.0.

The variable 0{.variable} [defined in modules/codec/avcodec/avcodec.h](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/avcodec/avcodec.h;h=5e526a3b1cd61d9eb90d79223994c115c1ac35e1;hb=HEAD#l230) does not mention [that it checks](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/avcodec/encoder.c;h=2f8e2d8a145c2558f57c97787eba2af407ec6af3;hb=HEAD#l477) for 1 (Low Delay) and 2 (Extended Low Delay). FIXME: Unclear whether they are actually supported.

Options as of 3.0.6:

#### Decoding

-   **avcodec-dr \** : Direct rendering *default value: enabled*
-   **avcodec-corrupted \** : Prefer visual artifacts instead of missing frames *default value: enabled*
-   **avcodec-error-resilience \** : libavcodec can do error resilience. However, with a buggy encoder (such as the ISO MPEG-4 encoder from M\$) this can produce a lot of errors. Valid values range from 0 to 4 (0 disables all errors resilience) *default value: 1*
-   **avcodec-workaround-bugs \** : Try to fix some bugs: 1 autodetect, 2 old msmpeg4, 4 xvid interlaced, 8 ump4, 16 no padding, 32 ac vlc, 64 Qpel chroma. This must be the sum of the values. For example, to fix "ac vlc" and "ump4", enter 40.) *default value: 1*
-   **avcodec-hurry-up \** : The decoder can partially decode or skip frame(s) when there is not enough time. It's useful with low CPU power but it can produce distorted pictures *default value: enabled*
-   **avcodec-skip-frame \ {-1,0,1,2,3,4}** : Force skipping of frames to speed up decoding (-1=None, 0=Default, 1=B-frames, 2=P-frames, 3=B+P frames, 4=all frames) *default value: 0*
-   **avcodec-skip-idct \ {-1,0,1,2,3,4}** : Force skipping of [IDCT](http://en.wikipedia.org/wiki/IDCT#DCT-III) to speed up decoding for frame types (-1=None, 0=Default, 1=B-frames, 2=P-frames, 3=B+P frames, 4=all frames) *default value: 0*
-   **avcodec-fast \** : Allow non specification compliant speedup tricks. Faster but error-prone *default value: disabled*
-   **avcodec-skiploopfilter \ {0 (None), 1 (Non-ref), 2 (Bidir), 3 (Non-key), 4 (All)}** : Skipping the loop filter (aka deblocking) usually has a detrimental effect on quality. However it provides a big speedup for high definition streams *default value: 0*
-   **avcodec-debug \** : Set FFmpeg debug mask *default value: 0*
-   **avcodec-codec \** : Internal libavcodec codec name *default value: NULL*
-   **avcodec-hw \ {any,vdpau_avcodec,vaapi,vaapi_drm,none}** : This allows hardware decoding when available *default value: any*
-   **avcodec-threads \** : Number of threads used for decoding, 0 meaning auto *default value: 0*
-   **avcodec-options \** : Advanced options, in the form 0 *default value: NULL*

#### Encoding

-   **sout-avcodec-codec \** : Internal libavcodec codec name *default value: NULL*
-   **sout-avcodec-hq \ {rd,bits,simple}** : Quality level for the encoding of motions vectors (this can slow down the encoding very much) *default value: rd*
-   **sout-avcodec-keyint \** : Number of frames that will be coded for one key frame *default value: 0*
-   **sout-avcodec-bframes \** : Number of B-frames that will be coded between two reference frames *default value: 0*
-   **sout-avcodec-hurry-up \** : The encoder can make on-the-fly quality tradeoffs if your CPU can't keep up with the encoding rate. It will disable trellis quantization, then the rate distortion of motion vectors (hq), and raise the noise reduction threshold to ease the encoder's task *default value: disabled*
-   **sout-avcodec-interlace \** : Enable dedicated

algorithms for interlaced frames *default value: disabled*

-   **sout-avcodec-interlace-me \** : Enable interlaced motion estimation algorithms. This requires more CPU *default value: enabled*
-   **sout-avcodec-vt \** : Video bitrate tolerance in kbit/s *default value: 0*
-   **sout-avcodec-pre-me \** : Enable the pre-motion estimation algorithm *default value: disabled*
-   **sout-avcodec-rc-buffer-size \** : Rate control buffer size (in kbytes). A bigger buffer will allow for better rate control, but will cause a delay in the stream *default value: 0*
-   **sout-avcodec-rc-buffer-aggressivity \** : Rate control buffer aggressiveness *default value: 1.0*
-   **sout-avcodec-i-quant-factor \** : Quantization factor of I-frames, compared with P-frames (for instance 1.0 =\> same qscale for I and P frames) *default value: 0*
-   **sout-avcodec-noise-reduction \** : Enable a simple noise reduction algorithm to lower the encoding length and bitrate, at the expense of lower quality frames *default value: 0*
-   **sout-avcodec-mpeg4-matrix \** : Use the MPEG-4 quantization matrix for MPEG-2 encoding. This generally yields a better looking picture, while still retaining the compatibility with standard MPEG-2 decoders *default value: disabled*
-   **sout-avcodec-qmin \** : Minimum video quantizer scale *default value: 0*
-   **sout-avcodec-qmax \** : Maximum video quantizer scale *default value: 0*
-   **sout-avcodec-trellis \** : Enable trellis quantization (rate distortion for block coefficients) *default value: disabled*
-   **sout-avcodec-qscale \** : A fixed video quantizer scale for VBR encoding (accepted values: 0.01 to 255.0) *default value: 3*
-   **sout-avcodec-strict \** : Force a strict standard compliance when encoding (accepted values: -2 to 2) *default value: 0*
-   **sout-avcodec-lumi-masking \** : Raise the quantizer for very bright macroblocks *default value: 0.0*
-   **sout-avcodec-dark-masking \** : Raise the quantizer for very dark macroblocks *default value: 0.0*
-   **sout-avcodec-p-masking \** : Raise the quantizer for macroblocks with a high temporal complexity *default value: 0.0*
-   **sout-avcodec-border-masking \** : Raise the quantizer for macroblocks at the border of the frame *default value: 0.0*
-   **sout-avcodec-luma-elim-threshold \** : Eliminates luminance blocks when the PSNR isn't much changed. The H.264 specification recommends -4 *default value: 0*
-   **sout-avcodec-chroma-elim-threshold \** : Eliminates chrominance blocks when the PSNR isn't much changed. The H.264 specification recommends 7 *default value: 0*
-   **sout-avcodec-aac-profile \** : Specify the AAC audio profile to use for encoding the audio bitstream. It takes the following options: main, low, ssr (not supported), ltp, hev1, hev2. hev1 and hev2 are currently supported only with libfdk-aac enabled libavcodec *default value: low*
-   **sout-avcodec-options \** : Advanced options, in the form 0 *default value: NULL*

#### Source code

-   [modules/codec/avcodec](https://git.videolan.org/?p=vlc.git;a=tree;f=modules/codec/avcodec;hb=HEAD) (folder)
-   [modules/codec/avcodec/avcodec.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/avcodec/avcodec.c) (main file)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### avformat {#modules-avformat}

See also: Documentation:Modules/avcodec

Muxing options are provided as a sub-module and internally depend on the variable 0{.variable}.

#### Demux

Module: avformat

**Type**: Access demux

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Avformat demuxer

**Shortcut(s)**: -

-   **avformat-format \** : Internal libavcodec format name *default value: NULL*
-   **avformat-options \** : Advanced options, in the form {opt=val,opt2=val2} *default value: NULL*

#### Mux

Module: avformat

**Type**: Muxer

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Avformat muxer

**Shortcut(s)**: -

-   **sout-avformat-mux \** : Force use of a specific avformat muxer *default value: NULL*
-   **sout-avformat-options \** : Advanced options, in the form {opt=val,opt2=val2} *default value: NULL*
-   **sout-avformat-reset-ts \** : The muxed content will start near a 0 timestamp *default value: disabled*

#### Source code

-   [modules/demux/avformat](https://git.videolan.org/?p=vlc.git;a=tree;f=modules/demux/avformat;hb=HEAD) (folder)
-   [modules/demux/avformat/avformat.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/avformat/avformat.c) (file)
-   [modules/codec/avcodec/avcommon.h](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/avcodec/avcommon.h) (header, defines 0{.variable} and 1{.variable} shown here)
-   [modules/demux/avformat/avformat.h](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/avformat/avformat.h) (header, defines 0{.variable} and 1{.variable} shown here)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### avi {#modules-avi}

#### Demux

Module: avi

**Type**: Access demux

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: AVI demuxer

**Shortcut(s)**: (none)

-   **avi-interleaved \** : Force interleaved method *default value: disabled*
-   **avi-index \ {0,1,2,3}** : Recreate a index for the AVI file. Use this if your AVI file is damaged or incomplete (not seekable). 0 ("Ask for action"), 1 ("Always fix"), 2 ("Never fix"), 3 ("Fix when necessary") *default value: 0*

#### Mux

Module: avi

**Type**: Muxer

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: AVI muxer

**Shortcut(s)**: 0

-   **sout-avi-artist \** : Artist *default value: NULL*
-   **sout-avi-date \** : Date *default value: NULL*
-   **sout-avi-genre \** : Genre *default value: NULL*
-   **sout-avi-copyright \** : Copyright *default value: NULL*
-   **sout-avi-comment \** : Comment *default value: NULL*
-   **sout-avi-name \** : Name *default value: NULL*
-   **sout-avi-subject \** : Subject *default value: NULL*
-   **sout-avi-encoder \** : Encoder *default value: "VLC Media Player - " 0{.variable}*
-   **sout-avi-keywords \** : Keywords *default value: NULL*

#### Source code

-   [modules/demux/avi/avi.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/avi/avi.c)
-   [modules/mux/avi.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/mux/avi.c)
-   [modules/demux/avi](https://git.videolan.org/?p=vlc.git;a=tree;f=modules/demux/avi;hb=HEAD)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### avio {#modules-avio}

See also: RTMP

The access output module is a sub-module.

#### Access

Module: avio

**Type**: Access

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: libavformat AVIO access

**Shortcut(s)**: 0, 1

Other shortcuts for this module are RTMP-related and reflect protocols: 0, 1, 2, 3, 4.

-   **avio-options \** : Advanced options, in the form 0 *default value: NULL*

##### Access output

Module: avio

**Type**: Access output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: libavformat AVIO access output

**Shortcut(s)**: 0, 1

-   **sout-avio-options \** : Advanced options, in the form 0 *default value: NULL*

#### Source code

-   [modules/access/avio.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/avio.c) - (main file)
-   [modules/access/avio.h](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/avio.h) - (contains module descriptor)
-   [modules/codec/avcodec/avcommon.h](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/avcodec/avcommon.h) - (contains text for module options)
-   [libavformat/avio.h](https://git.videolan.org/?p=ffmpeg.git;a=blob;f=libavformat/avio.h) (ffmpeg.git) - (called by modules/access/avio.c and modules/access/avio.h)
-   [libavformat/avformat.h](https://git.videolan.org/?p=ffmpeg.git;a=blob;f=libavformat/avformat.h) (ffmpeg.git) - (called by modules/access/avio.h)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### bandwidth {#modules-bandwidth}

Module: bandwidth

**Type**: Access filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: limit incoming bandwidth

**Shortcut(s)**: -

#### Options

-   **access-bandwidth \** : The bandwidth module will drop any data in excess of that many bytes per seconds *default value: 65536*

#### Example

    % vlc --access-filter bandwidth --access-bandwidth 131072

Will limit incoming data to 128 kBytes/second (128\*1024 Bytes/second).

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### bda {#modules-bda}

Module: bda

**Type**: Access

**First VLC version**: 0.9.0

**Last VLC version**: 1.1.?

**Operating system(s)**: Windows

**Description**: DirectShow [BDA](http://en.wikipedia.org/wiki/Broadcast_Driver_Architecture) input

**Shortcut(s)**: -

This module was superseded by dtv sometime between 1.1 and 2.0.

The [options for this module were conditional](http://en.wikipedia.org/wiki/C_preprocessor#Conditional_compilation) upon the presence of either macro 0 (Windows 32-bit) or 1 (Windows CE), indicating that the target was a Windows system:

    # if defined(WIN32) || defined(WINCE)
     // Condition: Windows
    # else
     // Condition: Other
    # endif

Shortcuts to this module were 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13 and 14.

#### Source code

-   [modules/access/bda](https://git.videolan.org/?p=vlc/vlc-1.1.git;a=tree;f=modules/access/bda;hb=HEAD) (vlc/vlc-1.1.git)
-   [modules/access/bda/bda.c](https://git.videolan.org/?p=vlc/vlc-1.1.git;a=blob;f=modules/access/bda/bda.c) (vlc/vlc-1.1.git)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### beos {#modules-beos}

Module: beos

**Type**: Interface

**First VLC version**: 0.5.0?

**Last VLC version**: 1.0.6

**Operating system(s)**: all

**Description**: BeOS standard API

**Shortcut(s)**: -

#### Options

-   **beos-dvdmenus \** : Use DVD Menus *default value: enabled*

#### Source code

-   [modules/gui/beos/BeOS.cpp](https://git.videolan.org/?p=vlc/vlc-1.0.git;a=blob;f=modules/gui/beos/BeOS.cpp) (vlc/vlc-1.0.git)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### blend {#modules-blend}

Module: blend

**Type**: Video filter

**First VLC version**: 0.8.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Blend one picture with alpha onto another picture

**Shortcut(s)**: -

FIXME: example usage?

The module description refers to itself as *blend2.cpp*; this is because it was [rewritten in C++](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=fec8f97c7f712483d047e961b31b5cbec921692a). The rewrite fixed [Bug #5477](https://trac.videolan.org/vlc/ticket/5477).

#### Options

None.

#### Source code

-   [modules/video_filter/blend.cpp](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/blend.cpp) - current
-   [modules/video_filter/blend.c](https://git.videolan.org/?p=vlc/vlc-1.1.git;a=blob;f=modules/video_filter/blend.c) (vlc/vlc-1.1.git) - old

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### bluescreen {#modules-bluescreen}

**Mosaic framework (How-To)Modules:** mosaic (mosaic-bridge • bridge-in • bridge-out) • alphamask • bluescreen

Module: bluescreen

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Change the video's alpha channel

**Shortcut(s)**: 0

This filter can be used in the mosaic framework to set a video's alpha channel (or transparency) based on a pixel's color. This is also known as [green screen](http://en.wikipedia.org/wiki/green_screen) or chroma key blending and can be used to create effects like on most weather channels.

#### Options

-   **bluescreen-u \** : U chroma component. *default value: 120*
-   **bluescreen-v \** : V chroma component. *default value: 90*
-   **bluescreen-ut \** : Tolerance of the bluescreen blender on color variations for the U plane. A value between 10 and 20 seems sensible. *default value: 17*
-   **bluescreen-vt \** : Tolerance of the bluescreen blender on color variations for the V plane. A value between 10 and 20 seems sensible. *default value: 17*

#### Example

    % vlc -vvv --vlm-conf mosaic.vlm --mosaic-keep-picture --sub-filter mosaic

And the vlm config:

    new channel0 broadcast enabled
    setup channel0 input rushfondvert.avi
    setup channel0 output #duplicate{dst=mosaic-bridge{chroma=YUVA,vfilter=bluescreen},select=video}

    new background broadcast enabled
    setup background input redefined-nintendo.mpg
    control background play

    control channel0 play

Have a look at [people.videolan.org/\~dionoea/bluescreen2.mpg (archived)](https://web.archive.org/web/20060819104251/0) for an example of the VLC bluescreen filter. The overlay video is [rushfondvert.avi (archived)](https://web.archive.org/web/20061205222657/1) and features someone with a mask in front of a green background. The bluescreen module's default values are adjusted to remove the background from this video. For other videos you should use your favorite color editing tool to find out the appropriate U and V values.

#### Another example

Tested with VLC 2.0.0

    new channel0 broadcast enabled
    setup channel0 input rushfondvert.avi
    setup channel0 output #duplicate{dst=mosaic-bridge{height=270,width=360,chroma=YUVA,vfilter=bluescreen},select=video}:display

    new background broadcast enabled
    setup background input file:///mire.jpg

    control background play
    control channel0 play

#### Source code

-   [modules/video_filter/bluescreen.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/bluescreen.c)

#### See also

-   YUV

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### bonjour {#modules-bonjour}

Past versions of VLC may refer to a different module as bonjour:
with VLC 3.0.0 the old bonjour.c module^(section\ link)^ was [renamed to avahi.c](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=55280fa62cb68b71767778c56250352b4840b69a) and [bonjour.m was created](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=1baae638b5759ff092c7977ab17185975f7e6524).

None of these modules have options.

#### Services discovery

Module: bonjour.m

**Type**: Services discovery

**First VLC version**: 3.0.0

**Last VLC version**: -

**Operating system(s)**: macOS, tvOS, iOS

**Description**: [Bonjour](http://en.wikipedia.org/wiki/Bonjour_(software) "wikipedia:Bonjour (software)") Network Discovery

**Shortcut(s)**: 0, 1

##### Renderer discovery

Module: bonjour.m

**Type**: Renderer discovery

**First VLC version**: 3.0.0

**Last VLC version**: -

**Operating system(s)**: macOS, tvOS, iOS

**Description**: [Bonjour](http://en.wikipedia.org/wiki/Bonjour_(software) "wikipedia:Bonjour (software)") Renderer Discovery

**Shortcut(s)**: 0, 1

The [introduction of this probe](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=d8203596f9e6a772fdaa4dd8c52ba77e49261406) contains a note:

    Add a bonjour renderer submodule to the bonjour service discovery
    module, so it can discover chromecast renderers (for now) and others
    in the future.
    There is still some work needed to make it detect chromecast
    capabilities correctly and to not hardcode it to chromecast.
    (See the TODO comment)

The TODO comment in the same commit:

    // TODO: Detect rendered capabilities and adapt to work with not just chromecast

#### Services discovery (avahi)

Module: avahi.c

**Type**: Services discovery

**First VLC version**: 0.8.4

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: [Zeroconf](http://en.wikipedia.org/wiki/Zeroconf) services

**Shortcut(s)**: 0, 1

##### Renderer discovery (avahi)

Module: avahi_renderer.c

**Type**: Renderer discovery

**First VLC version**: 1.0.4

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: [Avahi](http://en.wikipedia.org/wiki/Avahi_(software) "wikipedia:Avahi (software)") [Zeroconf](http://en.wikipedia.org/wiki/Zeroconf) renderer Discovery

**Shortcut(s)**: 0, 1

The [introduction of this probe](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=4037348c022a8937d8153e1a72c16a6085f01d15) was not announced; it was experimental. A stable version [is upcoming](https://git.videolan.org/?p=vlc.git;a=blob;f=NEWS;h=a90762a649d5bb2b8eba03a68163c10f459c6426;hb=HEAD#l63) (currently scheduled for 4.0.0-dev).

#### Source code

bonjour.m:

-   [modules/services_discovery/bonjour.m](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/services_discovery/bonjour.m)

bonjour.c/avahi.c:

-   [modules/services_discovery/bonjour.c](https://git.videolan.org/?p=vlc/vlc-0.8.git;a=blob;f=modules/services_discovery/bonjour.c) (vlc/vlc-0.8.git)
-   [modules/services_discovery/avahi.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/services_discovery/avahi.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### bridge / in {#modules-bridge-in}

**Mosaic framework (How-To)Modules:** mosaic (mosaic-bridge • bridge-in • bridge-out) • alphamask • bluescreen

Module: bridge-in

**Type**: Stream output

**First VLC version**: 0.8.2

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Get elementary streams from the bridge framework

**Shortcut(s)**: 0

This module gets all the elementary streams sent to the bridge framework. It is used when streaming a mosaic to attach the audio streams to the mosaic output.

#### Options

-   **sout-bridge-in-delay \** : Pictures coming from the picture video outputs will be delayed according to this value (in milliseconds, should be ≥ 100 ms). For high values, you will need to raise caching values. *default value: 0*
-   **sout-bridge-in-id-offset \** : Offset to add to the stream IDs specified in bridge-out to obtain the stream IDs bridge-in will register. *default value: 8192*
-   **sout-bridge-in-name \** : Name of this bridge-in instance. If you do not need more than one bridge-in at a time, you can discard this option. *default value: "default"*
-   **sout-bridge-in-placeholder \** : If set to true, the bridge will discard all input elementary streams except if it doesn't receive data from another bridge-in. This can be used to configure a placeholder stream when the real source breaks. Source and placeholder streams should have the same format. *default value: disabled*
-   **sout-bridge-in-placeholder-delay \** : Delay (in ms) before the placeholder kicks in. *default value: 200*
-   **sout-bridge-in-placeholder-switch-on-iframe \** : If enabled, switching between the placeholder and the normal stream will only occur on I-frames. This will remove artifacts on stream switching at the expense of a slightly longer delay, depending on the frequency of I-frames in the streams. *default value: enabled*

#### Source code

-   [modules/stream_out/bridge.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/stream_out/bridge.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### bridge / out {#modules-bridge-out}

**Mosaic framework (How-To)Modules:** mosaic (mosaic-bridge • bridge-in • bridge-out) • alphamask • bluescreen

Module: bridge-out

**Type**: Stream output

**First VLC version**: 0.8.2

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Send an elementary stream to the bridge framework

**Shortcut(s)**: 0

This module sends an elementary stream to the bridge framework. It is used when streaming a mosaic to send the audio stream to the mosaic output.

#### Options

-   **sout-bridge-out-id \** : Integer identifier for this elementary stream. This will be used to "find" this stream later. *default value: 0*
-   **sout-bridge-out-in-name \** : Name of the destination bridge-in. If you do not need more than one bridge-in at a time, you can discard this option. *default value: "default"*

#### Source code

-   [modules/stream_out/bridge.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/stream_out/bridge.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### caca {#modules-caca}

Module: caca

**Type**: Video output

**First VLC version**: 0.7.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Colored ASCII Art video output

**Shortcut(s)**: -

Output video with colored text instead of the image. Uses [libcaca](http://libcaca.zoy.org/).

#### Usage

    % vlc --vout caca somevideo.avi

#### Source code

-   [modules/video_output/caca.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_output/caca.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### cdda {#modules-cdda}

See also: CD

Module: cdda

**Type**: Access

**First VLC version**: ≤ 0.8

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Read a CD

**Shortcut(s)**: 0, 1

The option 0 is new as of [\[98dd4c30db57f88a92be16aa694f5d9fda08c15c\]](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=98dd4c30db57f88a92be16aa694f5d9fda08c15c) (2016). The option 1 (seems) to be deprecated as of [\[43bb27d91ce344eee93df3c956cd2513e3eecc3c\]](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=43bb27d91ce344eee93df3c956cd2513e3eecc3c) (2018).

The options 0 plays a particular track (like 1 does). 2 and 3 seem to be hints to VLC to skip [disk sectors](http://en.wikipedia.org/wiki/disk_sector) at the beginning or end.

-   **cd-audio \** : Audio CD device
-   **cdda-track \** : NULL *default value: 0*
-   **cdda-first-sector \** : NULL *default value: -1*
-   **cdda-last-sector \** : NULL *default value: -1*
-   **cddb-server \** : Address of the CDDB server to use *default value: freedb.videolan.org*
-   **cddb-port \** : CDDB Server port to use *default value: 80*

#### Source code

-   [modules/access/cdda.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/cdda.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### clone {#modules-clone}

Module: clone

**Type**: Video output splitter

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Clone the video output window

**Shortcut(s)**: 0

You can use this module to play the video in more than one window to test different video outputs or display the same video on multiple screens on the same computer.

#### Options

-   **clone-count \** : Number of video windows in which to clone the video. *default value: 2*
-   **clone-vout-list \** : You can use specific video output modules for the clones. Use a comma-separated list of modules. *default value: ""*

#### Examples

    $ vlc --video-splitter=clone --clone-count=2 video.ogv

#### Source code

-   [modules/video_splitter/clone.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_splitter/clone.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### colorthres {#modules-colorthres}

Module: colorthres

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Turn the picture black and white except for some colors

**Shortcut(s)**: -

This filters turns most of the picture black and white except for some colors. This could be called a "Schindler's List" effect.

#### Options

-   **colorthres-color \** : Colors similar to this will be kept, others will be grayscaled. *default value: 0xFF0000 (red)*

&nbsp;

-   **colorthres-saturationthres \** : Saturation threshold *default value: 20*

&nbsp;

-   **colorthres-similaritythres \** : Similarity threshold *default value: 15*

#### Example

    % vlc --video-filter colorthres somevideo.avi

#### See also

-   [Schindler's List, Red dress](http://en.wikipedia.org/wiki/Image:Schindlers_list_red_dress.JPG) on wikipedia

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### crop {#modules-crop}

*Were you looking for croppadd, the current module?*

Module: Crop

**Type**: Video filter

**First VLC version**: -

**Last VLC version**: 2.0.9

**Operating system(s)**: all

**Description**: Remove borders of the video and replace them with black borders

**Shortcut(s)**: -

#### Options

-   **crop-geometry \** : Set the geometry of the zone to crop. This is set as 0
-   **autocrop \** : Automatically detect black borders and crop them *default value: disabled*
-   **autocrop-ratio-max \** : Maximum image ratio. The crop plugin will never automatically crop to a higher ratio (ie, to a more "flat" image). The value is ×1000: 1333 means 4⁄3 *default value: 2405*
-   **crop-ratio \** : Force a ratio (0 for automatic). Value is ×1000: 1333 means 4⁄3 *default value: 0*
-   **autocrop-time \** : The number of consecutive images with the same detected ratio (different from the previously detected ratio) to consider that ratio changed and trigger recrop *default value: 25*
-   **autocrop-diff \** : The minimum difference in the number of detected black lines to consider that ratio changed and trigger recrop *default value: 16*
-   **autocrop-non-black-pixels \** : The maximum of non-black pixels in a line to consider that the line is black *default value: 3*
-   **autocrop-skip-percent \** : Percentage of the line to consider while checking for black lines. This allows skipping logos in black borders and crop them anyway *default value: 17*
-   **autocrop-luminance-threshold \** : Maximum luminance to consider a pixel as black (0-255)\* *default value: 40*

#### Note

This must be a typo. Despite the [claim of a range of 0](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/crop.c;h=2584780d98bb8527f9aacf540721a4ed5b852833;hb=c638a67c52980404d2aa6f6851b455743a898820#l97) in the help text for the 1 option, [the call](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/crop.c;h=2584780d98bb8527f9aacf540721a4ed5b852833;hb=c638a67c52980404d2aa6f6851b455743a898820#l129) to [2](https://git.videolan.org/?p=vlc.git;a=blob;f=include/vlc_configuration.h;h=bdbb11026492436f7f7297e096a8c62f8e899b68;hb=c638a67c52980404d2aa6f6851b455743a898820#l344) would have limited this to effectively 3.

#### Source code

-   [modules/video_filter/crop.c](https://git.videolan.org/?p=vlc/vlc-2.0.git;a=blob;f=modules/video_filter/crop.c) (vlc/vlc-2.0.git)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### croppadd {#modules-croppadd}

*For the former module, see Documentation:Modules/crop*

Module: cropadd

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Video cropping filter

**Shortcut(s)**: -

#### Options

##### Crop

-   **croppadd-croptop \<integer \[0 .. 0{.variable}\]\>** : Pixels to crop from top
-   **croppadd-cropbottom \<integer \[0 .. 0{.variable}\]\>** : Pixels to crop from bottom
-   **croppadd-cropleft \<integer \[0 .. 0{.variable}\]\>** : Pixels to crop from left
-   **croppadd-cropright \<integer \[0 .. 0{.variable}\]\>** : Pixels to crop from right

##### Padd

-   **croppadd-paddtop \<integer \[0 .. 0{.variable}\]\>** : Pixels to add to top
-   **croppadd-paddbottom \<integer \[0 .. 0{.variable}\]\>** : Pixels to add to bottom
-   **croppadd-paddleft \<integer \[0 .. 0{.variable}\]\>** : Pixels to add to left
-   **croppadd-paddright \<integer \[0 .. 0{.variable}\]\>** : Pixels to add to right

#### Source code

-   [modules/video_filter/croppadd.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/croppadd.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### daala {#modules-daala}

Module: daala

**Type**: Muxer

**First VLC version**: 3.0.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Daala video encoder

**Shortcut(s)**: -

Support for this module comes from the libdaala library.

#### Options

-   **sout-daala-quality \** : Enforce a quality between 0 (lossless) and 511 (worst) *default value: 10*
-   **sout-daala-keyint \** : Enforce a keyframe interval between 1 and 1000 *default value: 256*
-   **sout-daala-chroma-fmt \ {420,444}** : Picking chroma format will force a conversion of the video into that format. 0 means 12 and 3 means 45 *default value: 420*

#### Source code

-   [modules/codec/daala.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/daala.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### daap {#modules-daap}

This module provided compatibility in accessing iTunes shares. It was removed with the note:

    - Remove broken daap plugin (unmaintained wrt VLC API changes, relies on
     an unmaintained library, probably unable to read content from new itunes,
     ...). Implementations exist in rhythmbox, xmms2 and
     daap-sharp, we should see if a proper lib exists or if we could
     make one

The library it depended upon was [libopendaap](https://linux.die.net/man/3/libopendaap).

#### Services discovery

Module: daap

**Type**: Services discovery

**First VLC version**: 0.8.2

**Last VLC version**: 0.9.?

**Operating system(s)**: all

**Description**: [DAAP](http://en.wikipedia.org/wiki/Digital_Audio_Access_Protocol) shares

**Shortcut(s)**: (none)

##### Options

None.

##### Access

Module: daap

**Type**: Access

**First VLC version**: 0.8.2

**Last VLC version**: 0.9.?

**Operating system(s)**: all

**Description**: [DAAP](http://en.wikipedia.org/wiki/Digital_Audio_Access_Protocol) access

**Shortcut(s)**: (none)

###### Options

None.

#### Source code

-   [\[024fa1c48391bdcff9f3ca3f19f8ebb03a6db1f8\]](https://git.videolan.org/?p=vlc/vlc-0.8.git;a=commitdiff;h=024fa1c48391bdcff9f3ca3f19f8ebb03a6db1f8) (introduction)
-   [modules/services_discovery/daap.c](https://git.videolan.org/?p=vlc/vlc-0.8.git;a=blob;f=modules/services_discovery/daap.c) (vlc/vlc-0.8.git)
-   [\[0900f11014557ea895a290d2c1518d739f97a8b6\]](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=commitdiff;h=0900f11014557ea895a290d2c1518d739f97a8b6) (removal)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### dc1394 {#modules-dc1394}

**dc1394**

VLC uses this protocol (or access module) to read data from a device or network.
This protocol is handled by the **dc1394** module.

Module: dc1394

**Type**: Access demux

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: IIDC (DCAM) FireWire input module

**Shortcut(s)**: -

#### Introduction

This is an access module for [IIDC (firewire) cameras](http://en.wikipedia.org/wiki/IIDC). It uses libdc1394 version 1 libraries. Seems to have been written for an Apple iSight, as there is support for 320x240 and 640x480, but no other sizes. Only supported on Linux.

#### Installation

-   Before installation ensure that raw1394, dc1394, and all other necessary libraries for these. libdc1394_control.so.13 is needed either in /usr/lib or /usr/local/lib.
-   Ensure that the modules are loaded. (should happen automatically)
-   During the configure stage add: --enable-dc1394

#### Usage

Various examples of how to use the dc1394 module are shown below. You can try some new options yourself, but many have defaults and will work without.

    vlc dc1394:cameraindex=3:size=640x480:fps=30:brightness=100
    vlc dc1394:/dev/video1394/0:capture=raw1394
    vlc dc1394:/dev/video1394/0:adev=/dev/audio
    vlc dc1394:/dev/video1394/0:size=640x480:fps=30:brightness=200:adev=/dev/dsp:channel=2
    vlc dc1394:/dev/video1394/0:size=320x240:fps=15:brightness=200
       --sout='#transcode{vcodec=mp4v,vb=3000,keyint=30}:std{access=udp,mux=ts,url=192.168.150.79}'

#### Future Modifications

Unsure of plans to support the upcoming libdc1394 v2 library. As the library is designed to be configurable it would be ideal that the user can input any sizes, formats, and framerates (and other features) supported by the camera. Most cameras with larger frame sizes are more expensive cameras thus demand is likely not high, however, may be useful for some. Will create problems as it requires the user to be sure of the cameras supported features (however almost all support 640x480 and 320x240, so we will never exit poorly).

#### Source code

-   [modules/access/dc1394.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/dc1394.c) (access module)

### deinterlace {#modules-deinterlace}

Module: deinterlace

**Type**: Video output filter

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Deinterlacing video filter

**Shortcut(s)**: -

#### Options

-   **sout-deinterlace-mode \ {discard,blend,mean,bob,linear,x,yadif,yadif2x,phosphor,ivtc}** : Streaming deinterlace mode. Deinterlace method to use for streaming
-   **sout-deinterlace-phosphor-chroma \ {1,2,3,4}** : Phosphor chroma mode for 4:2:0 input. Choose handling for colours in those output frames that fall across input frame boundaries.
    -   Latest (1): take chroma from new (bright) field. Good for interlaced input, such as videos from a camcorder
    -   AltLine (2): take chroma line 1 from top field, line 2 from bottom field, etc. Default, good for NTSC telecined input (anime DVDs, etc.)
    -   Blend (3): average input field chromas. May distort the colours of the new (bright) field, too
    -   Upconvert (4): output in 4:2:2 format (independent chroma for each field). Best simulation, but requires more CPU and memory bandwidth *default value: 2*
-   **sout-deinterlace-phosphor-dimmer \ {1,2,3,4}** : Phosphor old field dimmer strength: 1 (Off), 2 (Low), 3 (Medium), 4 (High). This controls the strength of the darkening filter that simulates CRT TV phosphor light decay for the old field in the Phosphor framerate *default value: 2*

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### delay {#modules-delay}

Module: delay

**Type**: Stream output

**First VLC version**: 2.0.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Delay a stream

**Shortcut(s)**: 0

#### Options

-   **sout-delay-id \** : Specify an identifier integer for this elementary stream. *default value: 0*
-   **sout-delay-delay \** : Specify a delay (in [ms](http://en.wiktionary.org/wiki/ms)) for this elementary stream. Positive means delay and negative means advance. *default value: 0*

#### Examples

From the changelog: 0

#### Source code

-   [modules/stream_out/delay.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/stream_out/delay.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### description {#modules-description}

Module: stream_out_description

**Type**: Stream output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Gathers ES info

**Shortcut(s)**: -

The only shortcut for this module is 0. This module has no options.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### dirac {#modules-dirac}

*Were you looking for the current module, schroedinger?*

The dirac demuxer will be removed. It has already been [removed](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=eb8ab8df4b1483e1cc299d96b5c43b738ee03d25) in 4.0.0-dev.

#### Demux

Module: dirac

**Type**: Access demux

**First VLC version**: 1.0.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Dirac video demuxer

**Shortcut(s)**: 0

##### Options

-   **dirac-dts-offset \** : Value to adjust dts by *default value: 0*

#### Mux

Module: dirac

**Type**: Muxer

**First VLC version**: 0.8.2

**Last VLC version**: 2.1.6

**Operating system(s)**: all

**Description**: Dirac video encoder using [dirac-research library](https://sourceforge.net/projects/dirac/files/dirac-codec/)

**Shortcut(s)**: (none)

A few notes (of no real importance as this is a removed module):

-   the option 0 is, unusually, a float with an integer range
-   the option 0 is [dependent](http://en.wikipedia.org/wiki/Conditional_compilation) on Dirac research version of at least 1.0.1
-   the options 0 and 1 are disabled when set to 2. They have a code comment:
    /\* NB, unforunately vlc doesn't have a concept of 'don't care' \*/
-   the options 0 and 1 were originally aliases for 2 and 3 prior to 1.0.0

##### Options

-   **sout-dirac-quality \** : If bitrate=0, use this value for constant quality *default value: 5.5*
-   **sout-dirac-bitrate \<integer \[-1 .. 0{.variable}\]\>** : A value \> 0 (in kbps) enables constant bitrate mode *default value: -1*
-   **sout-dirac-lossless \** : Lossless coding ignores bitrate and quality settings, allowing for perfect reproduction of the original *default value: disabled*
-   **sout-dirac-prefilter \ {none,cwm,rectlp,diaglp}** : Enable adaptive prefiltering: The options correspond to "none", "Centre Weighted Median", "Rectangular Linear Phase", "Diagonal Linear Phase" *default value: diaglp*
-   **sout-dirac-prefilter-strength \** : Higher value implies more prefiltering *default value: 1*
-   **sout-dirac-chroma-fmt \ {420,422,444}** : Picking chroma format will force a conversion of the video into that format: "4:2:0", "4:2:2" or "4:4:4" *default value: 420*
-   **sout-dirac-l1-sep \<integer \[-1 .. 0{.variable}\]\>** : Distance between *P-frames* *default value: -1*
-   **sout-dirac-num-l1 \<integer \[-1 .. 0{.variable}\]\>** : Number of *P-frames* per GOP *default value: -1*
-   **sout-dirac-coding-mode \ {auto,progressive,field}** : Field coding is where interlaced fields are coded seperately as opposed to a pseudo-progressive frame (auto - let encoder decide based upon input (Best), progressive - force coding frame as single picture, field - force coding frame as seperate interlaced fields) *default value: auto*
-   **sout-dirac-mv-prec \ {1,1/2,1/4,1/8}** : Motion vector precision in pels *default value: 1/2*
-   **sout-dirac-mc-blk-width \<integer \[-1 .. 0{.variable}\]\>** : Width of motion compensation blocks *default value: -1*
-   **sout-dirac-mc-blk-height \<integer \[-1 .. 0{.variable}\]\>** : Height of motion compensation blocks *default value: -1*
-   **sout-dirac-mc-blk-overlap \** : Amount (%) that each motion block should be overlapped by its neighbours *default value: -1*
-   **sout-dirac-me-simple-search \** : (Not recommended) Perform a simple (non hierarchical block matching motion vector search with search range of ±0{.variable}, ±1{.variable}) *default value: ""*
-   **sout-dirac-me-combined \** : Use chroma as part of the motion estimation process *default value: enabled*
-   **sout-dirac-dwt-intra \** : Intra picture DWT filter *default value: -1*
-   **sout-dirac-dwt-inter \** : Inter picture DWT filter *default value: -1*
-   **sout-dirac-dwt-depth \** : Number of DWT iterations (Also known as DWT levels) *default value: -1*
-   **sout-dirac-noac \** : Disable arithmetic coding—Use variable length codes instead, useful for very high bitrates *default value: disabled*

###### Advanced options

-   **sout-dirac-mc-blk-xblen \<integer \[-1 .. 0{.variable}\]\>** : Total horizontal block length including overlaps *default value: -1*
-   **sout-dirac-mc-blk-yblen \<integer \[-1 .. 0{.variable}\]\>** : Total vertical block length including overlaps *default value: -1*
-   **sout-dirac-multi-quant \** : Enable multiple quantizers per subband (one per codeblock) *default value: -1*
-   **sout-dirac-spartition \** : Enable spatial partitioning *default value: -1*
-   **sout-dirac-cpd \<float \[-1 .. 0{.variable}\]\>** : cycles per degree *default value: -1*

#### Source code

-   [modules/demux/dirac.c](https://git.videolan.org/?p=vlc/vlc-3.0.git;a=blob;f=modules/demux/dirac.c) (vlc/vlc-3.0.git) (demux)
-   [modules/codec/dirac.c](https://git.videolan.org/?p=vlc/vlc-2.1.git;a=blob;f=modules/codec/dirac.c) (vlc/vlc-2.1.git) (mux)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### direct3d {#modules-direct3d}

Module: direct3d

**Type**: Video output

**First VLC version**: 0.8.6

**Last VLC version**: -

**Operating system(s)**: Windows

**Description**: Direct3D video output

**Shortcut(s)**: -

**Note:** This is default on windows Vista since VLC 0.8.6. It is likely to be default for all Windows versions starting with VLC 0.9.0.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### directfb {#modules-directfb}

See also: DirectFB Compile

Module: directfb

**Type**: Video output

**First VLC version**: -

**Last VLC version**: 2.2.8

**Operating system(s)**: Linux

**Description**: Direct framebuffer video output

**Shortcut(s)**: -

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### directory {#modules-directory}

Module: directory

**Type**: Access

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: recursively add files from a directory into the playlist

**Shortcut(s)**: -

#### Options

-   **recursive \** : Specify behavior when dealing with subdirectories. Can be "ignore" to ignore sub directories, "collapse" to add sub directories without expanding them and "expand" to add sub directories and expand them *default value: "expand"*
-   **ignore-filetypes \** : Comma seperated list of file extensions to ignore when adding directory items to the playlist *default value: "m3u,db,nfo,jpg,gif,sfv,txt,sub,idx,srt,cue"*

#### Example

Open a directory:

    % vlc directory:///path/to/dir

You can also use:

    % vlc /path/to/dir

or

    % vlc dir:///path/to/dir

or even

    % vlc file:///path/to/dir

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### directx aout {#modules-directx-aout}

Module: directx

**Type**: Audio output

**First VLC version**: 0.4

**Last VLC version**: -

**Operating system(s)**: Windows

**Description**: DirectSound audio output

**Shortcut(s)**: -

#### Introduction

This is the default audio output for Windows.

It uses the normal DirectSound API that is present since Windows 2000.

#### Options

##### Output device

This option allows you to select the audio output device listed.

You have to provide the number of the device.

##### Use float32 output

This option allows you to enable or disable the high-quality float32 audio output mode (which is not well supported by some soundcards).

##### Speaker configuration

This options is able to select speaker configuration you want to use.

It allows:

-   "Windows default"
-   "Mono", one audio channel only
-   "Stereo"
-   "Quad", 2 in front, 2 in back
-   "5.1"
-   "7.1"

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### directx vout {#modules-directx-vout}

Module: directx

**Type**: Video output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: Windows

**Description**: DirectDraw video output

**Shortcut(s)**: -

**Note:** This stopped being default on Windows Vista since VLC 0.8.6. It is likely to not be default for all Windows versions starting with VLC 0.9.0. See Direct3D for more.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### display {#modules-display}

Module: display

**Type**: Stream output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Display stream output

**Shortcut(s)**: 0

#### Options

-   **sout-display-audio \** : Enable audio rendering. *default value: enabled*
-   **sout-display-video \** : Enable video rendering. *default value: enabled*
-   **sout-display-delay \** : Introduces a delay (ms) in the display of the stream. *default value: 100*

#### Examples

Transcode and stream a file while displaying it locally.
This will display the transcoded version:

    % vlc somevideo.avi --sout "#transcode{vcodec=mp2v,vb=2048,acodec=mpga,ab=96}:duplicate{dst=std{access=udp,mux=ts{ttl=12},url=239.255.1.1},dst=display}"

This will display the original version:

    % vlc somevideo.avi --sout "#duplicate{dst='transcode{vcodec=mp2v,vb=2048,acodec=mpga,ab=96}:std{access=udp,mux=ts{ttl=12},url=239.255.1.1}',dst=display}"

#### Source code

-   [modules/stream_out/display.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/stream_out/display.c)

#### See also

-   Documentation:Modules/duplicate

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### distort {#modules-distort}

Module: distort

**Type**: Video filter

**First VLC version**: -

**Last VLC version**: 0.8.6

**Operating system(s)**: all

**Description**: Wave, ripple, gradient, psychedelic video filters

**Shortcut(s)**: -

This module was split into wave, ripple, gradient, psychedelic video filters between VLC 0.8.6i and 0.9.0.

#### Source code

-   [modules/video_filter/distort.c](https://git.videolan.org/?p=vlc/vlc-0.8.git;a=blob;f=modules/video_filter/distort.c) (vlc/vlc-0.8.git)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### dshow {#modules-dshow}

Module: dshow

**Type**: Access demux

**First VLC version**: 0.7.0

**Last VLC version**: -

**Operating system(s)**: Windows

**Description**: [DirectShow](http://en.wikipedia.org/wiki/DirectShow) input

**Shortcut(s)**: -

#### Options

-   **dshow-vdev \** : Name of the video device that will be used by the DirectShow plugin. If you don't specify anything, the default device will be used *default value: NULL*
-   **dshow-adev \** : Name of the audio device that will be used by the DirectShow plugin. If you don't specify anything, the default device will be used *default value: NULL*
-   **dshow-size \** : Size of the video that will be displayed by the DirectShow plugin. If you don't specify anything the default size for your device will be used. You can specify a standard size (cif, d1, ...) or \x\ *default value: NULL*
-   **dshow-aspect-ratio \** : Define input picture aspect ratio to use. Default is 4:3 *default value: 4:3*
-   **dshow-chroma \** : Force the DirectShow video input to use a specific chroma format (eg. I420 (default), RV24, etc.) *default value: NULL*
-   **dshow-fps \** : Force the DirectShow video input to use a specific frame rate (eg. 0 means default, 25, 29.97, 50, 59.94, etc.) *default value: 0.0f*
-   **dshow-config \** : Show the properties dialog of the selected device before starting the stream *default value: disabled*
-   **dshow-tuner \** : Show the tuner properties \[channel selection\] page *default value: disabled*
-   **dshow-tuner-channel \** : Set the TV channel the tuner will set to (0 means default) *default value: 0*
-   **dshow-tuner-country \** : Set the tuner country code that establishes the current channel-to-frequency mapping (0 means default) *default value: 0*
-   **dshow-tuner-input \ {0,1,2}** : Select the tuner input type (Default/Cable/Antenna) *default value: 0*
-   **dshow-video-input \** : Select the video input source, such as composite, s-video, or tuner. Since these settings are hardware-specific, you should find good settings in the "Device config" area, and use those numbers here. -1 means that settings will not be changed *default value: -1*
-   **dshow-video-output \** : Select the video output type. See the "video input" option *default value: -1*
-   **dshow-audio-input \** : Select the audio input source. See the "video input" option *default value: -1*
-   **dshow-audio-output \** : Select the audio output type. See the "video input" option *default value: -1*
-   **dshow-amtuner-mode \ {0,1,2,3,4}** : AM Tuner mode. Can be one of Default (0), TV (1), AM Radio (2), FM Radio (3) or DSS (4) *default value: 0{.variable}*
-   **dshow-audio-channels \** : Select audio input format with the given number of audio channels (if non 0) *default value: 0*
-   **dshow-audio-samplerate \** : Select audio input format with the given sample rate (if non 0) *default value: 0*
-   **dshow-audio-bitspersample \** : Select audio input format with the given bits/sample (if non 0) *default value: 0*

#### Source code

-   [modules/access/dshow/dshow.cpp](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/dshow/dshow.cpp)
-   [modules/access/dshow](https://git.videolan.org/?p=vlc.git;a=tree;f=modules/access/dshow;hb=HEAD)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### dtv {#modules-dtv}

See also: Documentation:Modules/dvb

For ease of navigation options are split into subpages: /Linux options and /Windows options.
The dtv module uses [conditional compilation](http://en.wikipedia.org/wiki/conditional_compilation) to determine which apply; separate modules are not used.

#### Linux

Module: dtv

**Type**: Access

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: Digital Television and Radio

**Shortcut(s)**: 0, 1, 2

Other shortcuts include:

-   cable: 0, 1, 2
-   satellite 0, 1, 2
-   terrestrial 0, 1, 2, 3

##### Options

*See Documentation:Modules/dtv/Linux options*

#### Windows

Module: dtv

**Type**: Access

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: Windows

**Description**: Digital Television and Radio

**Shortcut(s)**: 0, 1, 2

Other shortcuts include:

-   cable: 0, 1, 2
-   satellite 0, 1, 2
-   terrestrial 0, 1, 2, 3, 4

##### Options

*See Documentation:Modules/dtv/Windows options*

#### Source code

-   [modules/access/dtv/access.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/dtv/access.c)
-   [modules/access/dtv](https://git.videolan.org/?p=vlc.git;a=tree;f=modules/access/dtv;hb=HEAD)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### dtv / Linux options {#modules-dtv-linux-options}

#### Options (dummy section)

-   **dvb-adapter \** : If there is more than one digital broadcasting adapter, the adapter number must be selected. Numbering starts from zero. *default value: 0*
-   **dvb-device \** : If the adapter provides multiple independent tuner devices, the device number must be selected. Numbering starts from zero. *default value: 0*
-   **dvb-budget-mode \** : Only useful programs are normally demultiplexed from the transponder. This option will disable demultiplexing and receive all programs. *default value: false*
-   **dvb-frequency \** : TV channels are grouped by transponder (a.k.a. multiplex) on a given frequency. This is required to tune the receiver. *default value: 0*
-   **dvb-inversion \ { -1, 0, 1 }** : If the demodulator cannot detect spectral inversion correctly, it needs to be configured manually. *default value: -1*

##### Terrestrial reception parameters

-   **dvb-bandwidth \ { 0, 10, 8, 7, 6, 5, 2 }** : Bandwidth (MHz) *default value: 0*
-   **dvb-transmission \ { -1, 1, 2, 4, 8, 16, 32 }** : Transmission mode *default value: 0*
-   **dvb-guard \ { "1/128", "1/32", "1/16", "19/256", "1/8", "19/128", "1/4" }** : Guard interval *default value: ""*

##### DVB-T reception parameters

-   **dvb-code-rate-hp \ { "", "0", "1/2", "3/5", "2/3", "3/4", "4/5", "5/6", "6/7", "7/8", "8/9", "9/10" }** : The code rate for Forward Error Correction can be specified. *default value: ""*
-   **dvb-code-rate-lp \ { "", "0", "1/2", "3/5", "2/3", "3/4", "4/5", "5/6", "6/7", "7/8", "8/9", "9/10" }** : The code rate for Forward Error Correction can be specified. *default value: ""*
-   **dvb-hierarchy \ { -1, 0, 1, 2, 4 }** : Hierarchy mode *default value: -1*
-   **dvb-plp-id \** : DVB-T2 Physical Layer Pipe *default value: 0*

##### ISDB-T reception parameters

-   **dvb-a-modulation \ { "", "QAM", "16QAM", "32QAM", "64QAM", "128QAM", "256QAM", "8VSB", "16VSB", "QPSK", "DQPSK", "8PSK", "16APSK", "32APSK" }** : The digital signal can be modulated according with different constellations (depending on the delivery system). If the demodulator cannot detect the constellation automatically, it needs to be configured manually. *default value: NULL*
-   **dvb-a-fec \ { "", "0", "1/2", "3/5", "2/3", "3/4", "4/5", "5/6", "6/7", "7/8", "8/9", "9/10" }** : The code rate for Forward Error Correction can be specified. *default value: NULL*
-   **dvb-a-count \** : Layer A segments count *default value: 0*
-   **dvb-a-interleaving \** : Layer A time interleaving *default value: 0*
-   **dvb-b-modulation \ { "", "QAM", "16QAM", "32QAM", "64QAM", "128QAM", "256QAM", "8VSB", "16VSB", "QPSK", "DQPSK", "8PSK", "16APSK", "32APSK" }** : The digital signal can be modulated according with different constellations (depending on the delivery system). If the demodulator cannot detect the constellation automatically, it needs to be configured manually. *default value: NULL*
-   **dvb-b-fec \ { "", "0", "1/2", "3/5", "2/3", "3/4", "4/5", "5/6", "6/7", "7/8", "8/9", "9/10" }** : The code rate for Forward Error Correction can be specified. *default value: NULL*
-   **dvb-b-count \** : Layer B segments count *default value: 0*
-   **dvb-b-interleaving \** : Layer B time interleaving *default value: 0*
-   **dvb-c-modulation \ { "", "QAM", "16QAM", "32QAM", "64QAM", "128QAM", "256QAM", "8VSB", "16VSB", "QPSK", "DQPSK", "8PSK", "16APSK", "32APSK" }** : The digital signal can be modulated according with different constellations (depending on the delivery system). If the demodulator cannot detect the constellation automatically, it needs to be configured manually. *default value: NULL*
-   **dvb-c-fec \ { "", "0", "1/2", "3/5", "2/3", "3/4", "4/5", "5/6", "6/7", "7/8", "8/9", "9/10" }** : The code rate for Forward Error Correction can be specified. *default value: NULL*
-   **dvb-c-count \** : Layer C segments count *default value: 0*
-   **dvb-c-interleaving \** : Layer C time interleaving *default value: 0*

##### Cable and satellite reception parameters

-   **dvb-modulation \ { "", "QAM", "16QAM", "32QAM", "64QAM", "128QAM", "256QAM", "8VSB", "16VSB", "QPSK", "DQPSK", "8PSK", "16APSK", "32APSK" }** : The digital signal can be modulated according with different constellations (depending on the delivery system). If the demodulator cannot detect the constellation automatically, it needs to be configured manually. *default value: NULL*
-   **dvb-srate \** : The symbol rate must be specified manually for some systems, notably DVB-C, DVB-S and DVB-S2. *default value: 0*
-   **dvb-fec \ { "", "0", "1/2", "3/5", "2/3", "3/4", "4/5", "5/6", "6/7", "7/8", "8/9", "9/10" }** : The code rate for Forward Error Correction can be specified. *default value: ""*

##### DVB-S2 parameters

-   **dvb-stream \** : Stream identifier *default value: 0*
-   **dvb-pilot \ { -1, 0, 1 }** : Pilot *default value: -1*
-   **dvb-rolloff \ { -1, 35, 20, 25 }** : Roll-off factor *default value: -1*

##### ISDB-S parameters

-   **dvb-ts-id \** : Transport stream ID *default value: 0*

##### Satellite equipment control

-   **dvb-polarization \ { "", "V", "H", "R", "L" }** : To select the polarization of the transponder, a different voltage is normally applied to the low noise block-downconverter (LNB). *default value: ""*
-   **dvb-voltage \** : "" *default value: 13*
-   **dvb-high-voltage \** : If the cables between the satellilte low noise block-downconverter and the receiver are long, higher voltage may be required. Not all receivers support this. *default value: false*
-   **dvb-lnb-low \** : The downconverter (LNB) will subtract the local oscillator frequency from the satellite transmission frequency. The intermediate frequency (IF) on the RF cable is the result. *default value: 0*
-   **dvb-lnb-high \** : The downconverter (LNB) will subtract the local oscillator frequency from the satellite transmission frequency. The intermediate frequency (IF) on the RF cable is the result. *default value: 0*
-   **dvb-lnb-switch \** : If the satellite transmission frequency exceeds the switch frequency, the oscillator high frequency will be used as reference. Furthermore the automatic continuous 22kHz tone will be sent. *default value: 11700000*
-   **dvb-satno \ { 0, 1, 2, 3, 4 }** : If the satellite receiver is connected to multiple low noise block-downconverters (LNB) through a DiSEqC 1.0 switch, the correct LNB can be selected (1 to 4). If there is no switch, this parameter should be 0. *default value: 0*
-   **dvb-uncommitted \ { 0, 1, 2, 3, 4 }** : If the satellite receiver is connected to multiple low noise block-downconverters (LNB) through a cascade formed from DiSEqC 1.1 uncommitted switch and DiSEqC 1.0 committed switch, the correct uncommitted LNB can be selected (1 to 4). If there is no uncommitted switch, this parameter should be 0. *default value: 0*
-   **dvb-tone \ { -1, 0, 1 }** : A continuous tone at 22kHz can be sent on the cable. This normally selects the higher frequency band from a universal LNB. *default value: -1*

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### dtv / Windows options {#modules-dtv-windows-options}

#### Options (dummy section)

-   **dvb-adapter \** : If there is more than one digital broadcasting adapter, the adapter number must be selected. Numbering starts from zero. *default value: -1*
-   **dvb-network-name \** : Unique network name in the System Tuning Spaces *default value: ""*
-   **dvb-create-name \** : Create unique name in the System Tuning Spaces *default value: ""*
-   **dvb-frequency \** : TV channels are grouped by transponder (a.k.a. multiplex) on a given frequency. This is required to tune the receiver. *default value: 0*
-   **dvb-inversion \ { -1, 0, 1 }** : If the demodulator cannot detect spectral inversion correctly, it needs to be configured manually. *default value: -1*

##### Terrestrial reception parameters

-   **dvb-bandwidth \ { 0, 10, 8, 7, 6, 5, 2 }** : Bandwidth (MHz) *default value: 0*
-   **dvb-transmission \ { -1, 1, 2, 4, 8, 16, 32 }** : Transmission mode *default value: 0*
-   **dvb-guard \ { "1/128", "1/32", "1/16", "19/256", "1/8", "19/128", "1/4" }** : Guard interval *default value: ""*

##### DVB-T reception parameters

-   **dvb-code-rate-hp \ { "", "0", "1/2", "3/5", "2/3", "3/4", "4/5", "5/6", "6/7", "7/8", "8/9", "9/10" }** : The code rate for Forward Error Correction can be specified. *default value: ""*
-   **dvb-code-rate-lp \ { "", "0", "1/2", "3/5", "2/3", "3/4", "4/5", "5/6", "6/7", "7/8", "8/9", "9/10" }** : The code rate for Forward Error Correction can be specified. *default value: ""*
-   **dvb-hierarchy \ { -1, 0, 1, 2, 4 }** : Hierarchy mode *default value: -1*
-   **dvb-plp-id \** : DVB-T2 Physical Layer Pipe *default value: 0*

##### ISDB-T reception parameters

-   **dvb-a-modulation \ { "", "QAM", "16QAM", "32QAM", "64QAM", "128QAM", "256QAM", "8VSB", "16VSB", "QPSK", "DQPSK", "8PSK", "16APSK", "32APSK" }** : The digital signal can be modulated according with different constellations (depending on the delivery system). If the demodulator cannot detect the constellation automatically, it needs to be configured manually. *default value: NULL*
-   **dvb-a-fec \ { "", "0", "1/2", "3/5", "2/3", "3/4", "4/5", "5/6", "6/7", "7/8", "8/9", "9/10" }** : The code rate for Forward Error Correction can be specified. *default value: NULL*
-   **dvb-a-count \** : Layer A segments count *default value: 0*
-   **dvb-a-interleaving \** : Layer A time interleaving *default value: 0*
-   **dvb-b-modulation \ { "", "QAM", "16QAM", "32QAM", "64QAM", "128QAM", "256QAM", "8VSB", "16VSB", "QPSK", "DQPSK", "8PSK", "16APSK", "32APSK" }** : The digital signal can be modulated according with different constellations (depending on the delivery system). If the demodulator cannot detect the constellation automatically, it needs to be configured manually. *default value: NULL*
-   **dvb-b-fec \ { "", "0", "1/2", "3/5", "2/3", "3/4", "4/5", "5/6", "6/7", "7/8", "8/9", "9/10" }** : The code rate for Forward Error Correction can be specified. *default value: NULL*
-   **dvb-b-count \** : Layer B segments count *default value: 0*
-   **dvb-b-interleaving \** : Layer B time interleaving *default value: 0*
-   **dvb-c-modulation \ { "", "QAM", "16QAM", "32QAM", "64QAM", "128QAM", "256QAM", "8VSB", "16VSB", "QPSK", "DQPSK", "8PSK", "16APSK", "32APSK" }** : The digital signal can be modulated according with different constellations (depending on the delivery system). If the demodulator cannot detect the constellation automatically, it needs to be configured manually. *default value: NULL*
-   **dvb-c-fec \ { "", "0", "1/2", "3/5", "2/3", "3/4", "4/5", "5/6", "6/7", "7/8", "8/9", "9/10" }** : The code rate for Forward Error Correction can be specified. *default value: NULL*
-   **dvb-c-count \** : Layer C segments count *default value: 0*
-   **dvb-c-interleaving \** : Layer C time interleaving *default value: 0*

##### Cable and satellite reception parameters

-   **dvb-modulation \ { "", "QAM", "16QAM", "32QAM", "64QAM", "128QAM", "256QAM", "8VSB", "16VSB", "QPSK", "DQPSK", "8PSK", "16APSK", "32APSK" }** : The digital signal can be modulated according with different constellations (depending on the delivery system). If the demodulator cannot detect the constellation automatically, it needs to be configured manually. *default value: NULL*
-   **dvb-srate \** : The symbol rate must be specified manually for some systems, notably DVB-C, DVB-S and DVB-S2. *default value: 0*
-   **dvb-fec \ { "", "0", "1/2", "3/5", "2/3", "3/4", "4/5", "5/6", "6/7", "7/8", "8/9", "9/10" }** : The code rate for Forward Error Correction can be specified. *default value: ""*

##### DVB-S2 parameters

-   **dvb-stream \** : Stream identifier *default value: 0*
-   **dvb-pilot \ { -1, 0, 1 }** : Pilot *default value: -1*
-   **dvb-rolloff \ { -1, 35, 20, 25 }** : Roll-off factor *default value: -1*

##### ISDB-S parameters

-   **dvb-ts-id \** : Transport stream ID *default value: 0*

##### Satellite equipment control

-   **dvb-polarization \ { "", "V", "H", "R", "L" }** : To select the polarization of the transponder, a different voltage is normally applied to the low noise block-downconverter (LNB). *default value: ""*
-   **dvb-voltage \** : "" *default value: 13*
-   **dvb-lnb-low \** : The downconverter (LNB) will subtract the local oscillator frequency from the satellite transmission frequency. The intermediate frequency (IF) on the RF cable is the result. *default value: 0*
-   **dvb-lnb-high \** : The downconverter (LNB) will subtract the local oscillator frequency from the satellite transmission frequency. The intermediate frequency (IF) on the RF cable is the result. *default value: 0*
-   **dvb-lnb-switch \** : If the satellite transmission frequency exceeds the switch frequency, the oscillator high frequency will be used as reference. Furthermore the automatic continuous 22kHz tone will be sent. *default value: 11700000*
-   **dvb-network-id \** : Network identifier *default value: 0*
-   **dvb-azimuth \** : Satellite azimuth in tenths of degree *default value: 0*
-   **dvb-elevation \** : Satellite elevation in tenths of degree *default value: 0*
-   **dvb-longitude \** : Satellite longitude in tenths of degree. West is negative. *default value: 0*
-   **dvb-range \** : Satellite range code as defined by manufacturer e.g. DISEqC switch code *default value: ""*

##### ATSC reception parameters

-   **dvb-major-channel \** : Major channel *default value: 0*
-   **dvb-minor-channel \** : ATSC minor channel *default value: 0*
-   **dvb-physical-channel \** : Physical channel *default value: 0*

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### dummy {#modules-dummy}

Dummy modules are the VLC equivalent of 0 on GNU/Linux: they represent doing nothing.

#### List

-   **Type**: Description Shortcut(s)
-   **Access**: Dummy input dummy, vlc
-   **Access output**: Dummy stream output dummy
-   **Audio output**: Dummy audio output dummy
-   **Decoder**: Dummy decoder dummy, dump
-   **Encoder**: Dummy encoder dummy
-   **Control interface**: Dummy interface (none)
-   **Muxer**: Dummy/Raw muxer dummy, raw, es
-   **Stream output**: Dummy stream output dummy, drop
-   **Text renderer**: Dummy font renderer (none)
-   **Video output**: Dummy video output dummy, stats
-   **Video output (legacy video plugins)**: Dummy window dummy

#### Descriptions

##### Interface

A dummy interface works well with command-line usage. It consumes fewer computing resources, something that can be used to reduce a bottleneck effect during transcoding, or simply when opening a window makes little sense.

This will play an Ogg file with no interface:

    $ vlc -I dummy audio.ogg vlc://quit

No window is launched, and control is returned to the calling application after the file plays.

This will play a Schroedinger file with a minimalist interface:

    $ vlc -I dummy video.drc vlc://quit

A window with no buttons or toolbars is launched (no skin), because video output requires a window. Hotkeys may be used to control playback by default, something that can be disabled if necessary with 0.

##### Stream output

Doesn't do anything. Can be used to test other stream output chain modules without actually streaming anything.

#### Options

##### Dummy decoder

-   **dummy-save-es \** : Save the raw codec data if you have selected/forced the dummy decoder in the main options. *default value: disabled*

##### Video output

-   **dummy-chroma \** : Force the dummy video output to create images using a specific chroma format instead of trying to improve performances by using the most efficient one. *default value: NULL*

#### Source code

-   [modules/access/idummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/idummy.c) (Access)
-   [modules/access_output/dummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access_output/dummy.c) (Access output)
-   [modules/audio_output/adummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/audio_output/adummy.c) (Audio output)
-   [modules/codec/ddummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/ddummy.c) (Decoder)
-   [modules/codec/edummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/edummy.c) (Encoder)
-   [modules/control/dummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/control/dummy.c) (Interface)
-   [modules/mux/dummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/mux/dummy.c) (Output muxer)
-   [modules/stream_out/dummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/stream_out/dummy.c) (Stream output)
-   [modules/text_renderer/tdummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/text_renderer/tdummy.c) (Text rendering)
-   [modules/video_output/vdummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_output/vdummy.c) (Video output)
-   [modules/video_output/wdummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_output/wdummy.c) (Video output for legacy video plugins)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### dummy sout {#modules-dummy-sout}

Dummy modules are the VLC equivalent of 0 on GNU/Linux: they represent doing nothing.

#### List

-   **Type**: Description Shortcut(s)
-   **Access**: Dummy input dummy, vlc
-   **Access output**: Dummy stream output dummy
-   **Audio output**: Dummy audio output dummy
-   **Decoder**: Dummy decoder dummy, dump
-   **Encoder**: Dummy encoder dummy
-   **Control interface**: Dummy interface (none)
-   **Muxer**: Dummy/Raw muxer dummy, raw, es
-   **Stream output**: Dummy stream output dummy, drop
-   **Text renderer**: Dummy font renderer (none)
-   **Video output**: Dummy video output dummy, stats
-   **Video output (legacy video plugins)**: Dummy window dummy

#### Descriptions

##### Interface

A dummy interface works well with command-line usage. It consumes fewer computing resources, something that can be used to reduce a bottleneck effect during transcoding, or simply when opening a window makes little sense.

This will play an Ogg file with no interface:

    $ vlc -I dummy audio.ogg vlc://quit

No window is launched, and control is returned to the calling application after the file plays.

This will play a Schroedinger file with a minimalist interface:

    $ vlc -I dummy video.drc vlc://quit

A window with no buttons or toolbars is launched (no skin), because video output requires a window. Hotkeys may be used to control playback by default, something that can be disabled if necessary with 0.

##### Stream output

Doesn't do anything. Can be used to test other stream output chain modules without actually streaming anything.

#### Options

##### Dummy decoder

-   **dummy-save-es \** : Save the raw codec data if you have selected/forced the dummy decoder in the main options. *default value: disabled*

##### Video output

-   **dummy-chroma \** : Force the dummy video output to create images using a specific chroma format instead of trying to improve performances by using the most efficient one. *default value: NULL*

#### Source code

-   [modules/access/idummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/idummy.c) (Access)
-   [modules/access_output/dummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access_output/dummy.c) (Access output)
-   [modules/audio_output/adummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/audio_output/adummy.c) (Audio output)
-   [modules/codec/ddummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/ddummy.c) (Decoder)
-   [modules/codec/edummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/edummy.c) (Encoder)
-   [modules/control/dummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/control/dummy.c) (Interface)
-   [modules/mux/dummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/mux/dummy.c) (Output muxer)
-   [modules/stream_out/dummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/stream_out/dummy.c) (Stream output)
-   [modules/text_renderer/tdummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/text_renderer/tdummy.c) (Text rendering)
-   [modules/video_output/vdummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_output/vdummy.c) (Video output)
-   [modules/video_output/wdummy.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_output/wdummy.c) (Video output for legacy video plugins)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### dump {#modules-dump}

Module: dump

**Type**: Access filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: dump data to disk

**Shortcut(s)**: -

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### duplicate {#modules-duplicate}

Module: stream_out_duplicate

**Type**: Stream output

**First VLC version**: 1.1.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Duplicate stream output

**Shortcut(s)**: 0, 1

#### Options

None.

#### Examples

From the changelog: 0{.nowrap}

#### Source code

-   [modules/stream_out/duplicate.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/stream_out/duplicate.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### dvb {#modules-dvb}

See also: Documentation:Modules/dtv

Module: dvb

**Type**: Access

**First VLC version**: 0.6.2

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: DVB input with v4l2 support

**Shortcut(s)**: 0

Shortcuts:

-   0 (Generic name)
    -   0, 1, 2 (Satellite)
    -   0, 1 (Cable)
    -   0, 1 (Terrestrial)

#### Options

-   **dvb-probe \** : Some DVB cards do not like to be probed for their capabilities, you can disable this feature if you experience some trouble. *default value: enabled*
-   **dvb-satellite \** : Filename of config file in share/dvb/dvb-s. *default value: NULL*
-   **dvb-scanlist \** : Filename containing initial scan tuning data. *default value: NULL*
-   **dvb-scan-nit \** : Use NIT for scanning services *default value: enabled*

#### See also

-   dvbsub

#### Source code

-   [modules/access/dvb/access.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/dvb/access.c) (main file)
-   [modules/access/dvb](https://git.videolan.org/?p=vlc.git;a=tree;f=modules/access/dvb;hb=HEAD) (folder)
-   [modules/demux/playlist/dvb.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/playlist/dvb.c) (LinuxTV channels list, part of playlist)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### dvbsub {#modules-dvbsub}

Module: dvbsub

**Type**: Subtitles

**First VLC version**: 0.8.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: DVB subtitles decoder/encoder

**Shortcut(s)**: -

There are notes in the source code:

    /*
     * Preamble
     *
     * FIXME:
     * DVB subtitles coded as strings of characters are not handled correctly.
     * The character codes in the string should actually be indexes referring to a
     * character table identified in the subtitle descriptor.
     *
     * The spec is quite vague in this area, but what is meant is perhaps that it
     * refers to the character index in the codepage belonging to the language
     * specified in the subtitle descriptor. Potentially it's designed for widechar
     * (but not for UTF-*) codepages.
     */

and

    /*
     * Notes on DDS (Display Definition Segment)
     * -----------------------------------------
     * DDS (Display Definition Segment) tells the decoder how the subtitle image
     * relates to the video image.
     * For SD, the subtitle image is always considered to be for display at
     * 720x576 (although it's assumed that for NTSC, this is 720x480, this
     * is not documented well) Also, for SD, the subtitle image is drawn 'on
     * the glass' (i.e. after video scaling, letterbox, etc.)
     * For 'HD' (subs marked type 0x14/0x24 in PSI), a DDS must be present,
     * and the subs area is drawn onto the video area (scales if necessary).
     * The DDS tells the decoder what resolution the subtitle images were
     * intended for, and hence how to scale the subtitle images for a
     * particular video size
     * i.e. if HD video is presented as letterbox, the subs will be in the
     * same place on the video as if the video was presented on an HD set
     * indeed, if the HD video was pillarboxed by the decoder, the subs may
     * be cut off as well as the video. The intent here is that the subs can
     * be placed accurately on the video - something which was missed in the
     * original spec.
     *
     * A DDS may also specify a window - this is where the subs images are moved so that the (0,0)
     * origin of decode is offset.
     /

#### Options

-   **dvbsub-position \ { 0, 1, 2, 4, 8, 5, 6, 9, 10 }** : Subpicture position^(**key**)^
-   **dvbsub-x \** : X coordinate of the rendered subtitle *default value: -1*
-   **dvbsub-y \** : Y coordinate of the rendered subtitle *default value: -1*

##### Encoder

-   **sout-dvbsub-x \** : X coordinate of the encoded subtitle *default value: -1*
-   **sout-dvbsub-y \** : Y coordinate of the encoded subtitle *default value: -1*

#### Appendix

-   \^ --dvbsub-position

&nbsp;

-   **Integer**: Alignment Comment
-   **0**: Center
-   **1**: Left
-   **2**: Right
-   **4**: Top
-   **8**: Bottom
-   **5**: Top-Left 4 + 1
-   **6**: Top-Right 4 + 2
-   **9**: Bottom-Left 8 + 1
-   **10**: Bottom-Right 8 + 2
-   **3**: n/a contradictory
-   **7**: n/a contradictory

#### Source code

-   [modules/codec/dvbsub.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/dvbsub.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### dvdnav {#modules-dvdnav}

See also: Documentation:Modules/dvdread

Module: dvdnav

**Type**: Access demux

**First VLC version**: 0.7.1

**Last VLC version**: -

**Operating system(s)**: all

**Description**: DVD with menus

**Shortcut(s)**: -

This module uses libdvdnav.

#### Options

-   **dvdnav-angle \** : Default DVD angle *default value: 1*
-   **dvdnav-menu \** : Start the DVD directly in the main menu. This will try to skip all the useless warning introductions *default value: enabled*

#### Source code

-   [modules/access/dvdnav.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/dvdnav.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### dvdread {#modules-dvdread}

See also: Documentation:Modules/dvdnav

Module: dvdread

**Type**: Access demux

**First VLC version**: 0.8.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: DVD without menus

**Shortcut(s)**: -

Shortcuts to this module include 0 and 1. This module uses libdvdread. The option 2 has been deprecated since 1.1.0.

#### Options

-   **dvdread-angle \** : Default DVD angle *default value: 1*

#### Source code

-   [modules/access/dvdread.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/dvdread.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### erase {#modules-erase}

See also: Documentation:Modules/logo

Module: erase

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: logo erasing video filter

**Shortcut(s)**: -

Use this filter to erase a logo (or any given area) from the video.

#### Options

-   **erase-mask \** : PNG file to use as a mask. The alpha channel only will be used to build the mask *default value: NULL*
-   **erase-x \** : X offset from upper left corner *default value: 0*
-   **erase-y \** : Y offset from upper left corner. *default value: 0*

#### Example

    $ vlc --video-filter "erase{mask=logo.png,x=100,y=50}" somevideo.avi

#### Source code

-   [modules/video_filter/erase.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/erase.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### esd {#modules-esd}

Module: esd

**Type**: Audio output

**First VLC version**: 0.0.95

**Last VLC version**: 0.9.10

**Operating system(s)**: Linux

**Description**: [EsounD](http://en.wikipedia.org/wiki/EsounD) audio output

**Shortcut(s)**: -

Enlightened Sound Daemon (also known as EsounD and ESD) was one of the original audio output modules for Linux included since VLC's infancy. arts and the esd plugin were removed prior to VLC 1.0.0, because both projects were inactive.

Modern Linux users can use the pulse or jack modules instead. There are probably others.

#### Options

None.

#### Source code

-   [modules/audio_output/esd.c](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=blob;f=modules/audio_output/esd.c) (vlc/vlc-0.9.git)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### extract {#modules-extract}

Module: extract

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: extract a color component from the video

**Shortcut(s)**: -

-   **extract-component \** : Color component to extract (0xRRGGBB) *default value: 0xFF0000 (red)*

#### Example

Extract the yellow component from a video:

    % vlc --video-filter "extract{component=0xFFFF00}" somevideo.avi

#### Typical use

You can create a live Andy Warhol display.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### eyetv {#modules-eyetv}

Module: eyetv

**Type**: Access

**First VLC version**: 0.9.0

**Last VLC version**: 2.2.8

**Operating system(s)**: macOS

**Description**: [EyeTV](http://en.wikipedia.org/wiki/EyeTV) input

**Shortcut(s)**: -

#### Options

-   **eyetv-channel \** : EyeTV program number, or use 0 for last channel, -1 for S-Video input, -2 for Composite input *default value: 0*

#### Source code

-   [modules/access/eyetv.m](https://git.videolan.org/?p=vlc/vlc-2.2.git;a=blob;f=modules/access/eyetv.m) (vlc/vlc-2.2.git)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### faad {#modules-faad}

See also: FAAD2

Module: faad

**Type**: Audio decoder

**First VLC version**: 0.5.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: AAC audio decoder (using libfaad2)

**Shortcut(s)**: (none)

#### Options

None.

#### Source code

-   [modules/codec/faad.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/faad.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### fake {#modules-fake}

#### Options

##### Access Demux

Module: fake

**Type**: Access demux

**First VLC version**: -

**Last VLC version**: 0.9.0

**Operating system(s)**: all

**Description**: simulate a fake input

**Shortcut(s)**: 0

-   **fake-caching \** : Caching in [milliseconds](http://en.wiktionary.org/wiki/ms) *default value: 0{.variable}1*
-   **fake-fps \** : Framerate e.g. 24, 25, 29.97, 30 *default value: 25.0*
-   **fake-id \** : Set the ID of the fake elementary stream for use in 0{.sample}1{.sample}2{.sample} constructs *default value: 0*
-   **fake-duration \** : Duration of the fake streaming (in milliseconds) before faking an end-of-file (default is 0, meaning that the stream is unlimited) *default value: 0*

##### Codec

Module: fake

**Type**: Codec

**First VLC version**: -

**Last VLC version**: 0.9.0

**Operating system(s)**: all

**Description**: handle a fake input stream

**Shortcut(s)**: 0

-   **fake-file \** : Image to use as video for the fake stream *default value: ""*
-   **fake-file-reload \** : Number of seconds between each reload of the image *default value: 0*
-   **fake-width \** : Width *default value: 0*
-   **fake-height \** : Height *default value: 0*
-   **fake-keep-ar \** : Keep aspect ratio when resizing *default value: disabled*
-   **fake-aspect-ratio \** : Aspect ratio of the image file (4:3, 16:9). Default is square pixels *default value: ""*
-   **fake-deinterlace \** : Deinterlace the image after loading it *default value: disabled*
-   **fake-deinterlace-module \** : Deinterlace module *default value: "deinterlace"*
-   **fake-chroma \** : Image chroma *default value: "I420"*

#### Example

    $ vlc fake:// --fake-file someimage.png

#### Source code

-   [modules/access/fake.c](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=blob;f=modules/access/fake.c) (vlc/vlc-0.9.git)
-   [modules/codec/fake.c](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=blob;f=modules/codec/fake.c) (vlc/vlc-0.9.git)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### fdkaac {#modules-fdkaac}

Module: fdkaac

**Type**: Muxer

**First VLC version**: 2.1.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: FDK-AAC Audio encoder

**Shortcut(s)**: 0

This module is dual-licenced under LGPL 2.1 and BSD 2-clause. [modules/codec/avcodec.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/avcodec.c) (FAAC) used to handle encoding AAC, but it is [not used anymore](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=34ba8bd409b16a33353b2240330b405c970b0f7c).

#### Options

-   **sout-fdkaac-profile \ {2,5,29,23,39}** : Encoder Algorithm to use (2: AAC-LC, 5: HE-AAC, 29: HE-AAC-v2, 23: AAC-LD, 39: AAC-ELD) *default value: 0{.variable}*
-   **sout-fdkaac-sbr \** : Enable [spectral band replication](http://en.wikipedia.org/wiki/spectral_band_replication)—This is an optional feature only for the AAC-ELD profile *default value: disabled*
-   **sout-fdkaac-vbr \** : Quality of the VBR Encoding (0=cbr, 1-5 constant vbr quality, 5 is the best) *default value: 0*
-   **sout-fdkaac-afterburner \** : This library will produce higher quality audio at the expense of additional CPU usage (default is enabled) *default value: enabled*
-   **sout-fdkaac-signaling \** : 1 is explicit for SBR (0{.variable}) and implicit for PS (default), 2 is explicit hierarchical *default value: 1{.variable}*

#### Source code

-   [modules/codec/fdkaac.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/fdkaac.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### file {#modules-file}

Module: file

**Type**: Access

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: read a file

**Shortcut(s)**: -

#### Options

-   **file-caching \** : Caching value in ms

#### Example

Open a file:

    % vlc file:///path/to/somevideo.avi

The *file://* part can be left out unless the file name is ambiguous.

Open a non seekable stream:

    % vlc stream:///path/to/somepipe

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### file aout {#modules-file-aout}

Module: afile

**Type**: Audio output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Audio output to write to a file

**Shortcut(s)**: -

Shortcuts for this module include 0 and 1.

#### Options

-   **audiofile-file \** : File to which the audio samples will be written to ("-" for stdout) *default value: audiofile.wav*
-   **audiofile-format {u8,s16,float32,spdif}** : Output format *default value: s16*
-   **audiofile-channels \** : By default (0), all the channels of the incoming will be saved but you can restrict the number of channels here *default value: 0*
-   **audiofile-wav \** : Instead of writing a raw file, you can add a WAV header to the file *default value: enabled*

#### Source code

-   [modules/audio_output/file.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/audio_output/file.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### flac {#modules-flac}

FLAC is supported through libflac.

#### Decoder

Module: flac

**Type**: Audio decoder

**First VLC version**: 0.5.2

**Last VLC version**: -

**Operating system(s)**: all

**Description**: FLAC audio decoder

**Shortcut(s)**: 0

##### Options

None.

#### Encoder

Module: flac

**Type**: Audio encoder

**First VLC version**: 0.7.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: FLAC audio encoder

**Shortcut(s)**: 0

##### Options

None.

#### Demux

Module: flac

**Type**: Access demux

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: FLAC demuxer

**Shortcut(s)**: 0

##### Options

None.

#### Packetizer

Module: flac

**Type**: Packetizer

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Flac audio packetizer

**Shortcut(s)**: (none)

##### Options

None.

#### Source code

-   [modules/codec/flac.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/flac.c)
-   [modules/demux/flac.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/flac.c)
-   [modules/packetizer/flac.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/packetizer/flac.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### fluidsynth {#modules-fluidsynth}

[Wikipedia](http://en.wikipedia.org/wiki/Main_Page) has information on this entry:

***[Midi files](http://en.wikipedia.org/wiki/Musical_Instrument_Digital_Interface#Standard_MIDI_files)***

Module: smf

**Type**: Demux

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Standard MIDI Files

**Shortcut(s)**: -

Standard MIDI Files (SMF) contain sounds events that indicate the notes and instruments in a musical performance, but do not include the digital waveform of the audio. They usually have the extension 0 or 1. To play a MIDI file, software has to synthesize the music, which usually requires reading digital samples of musical instruments from a large file.

#### Play .mid (MIDI) files in VLC

Module: fluidsynth

**Type**: Codec

**First VLC version**: 0.9.0 (Linux)
1.1.0 (Windows)

**Last VLC version**: 3.0.x (Windows)

**Operating system(s)**: Linux

**Description**: MIDI synthesis with the FluidSynth library

**Shortcut(s)**: -

VLC media player can play Standard MIDI File (.MID) and RIFF MIDI (.RMI) files since version 0.9.0.

Windows binary builds included MIDI support only in versions VLC media player from 1.1.0 through 2.0.8. Starting from version 2.1.0, support was dropped due to [security issues](https://trac.videolan.org/vlc/ticket/9486). It was re-activated in VLC 3.0.0.

##### SoundFonts file

To playback MIDI files, you need a [SoundFont](http://en.wikipedia.org/wiki/SoundFont) file (with extension 0). You can download them from either of these two places:

-   0
-   0

##### Configure SoundFont in VLC

You need to open VLC's preferences. The preferences window has two display modes called **Simple** and **All**. Choose the display mode called **All**, then go to **Input/Codecs** \> **Audio codecs** \> **FluidSynth**. Then select the .sf2 file with **Browse** button and save the preferences with **Save** button.

##### Linux

If the **FluidSynth** codec is not shown in VLC's preferences, you have to install it as well as sound fonts. E.g. on Ubuntu 18.04 and derivatives it is in the **vlc-plugin-fluidsynth** package, while the **fluid-soundfont-gs** and **fluid-soundfont-gm** packages install some sound fonts in 0.

### freeze {#modules-freeze}

Module: freeze

**Type**: Video filter

**First VLC version**: 2.2.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Freezing interactive video filter

**Shortcut(s)**: -

#### Options

None.

#### Source code

-   [modules/video_filter/freeze.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/freeze.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### ftp {#modules-ftp}

Module: ftp

**Type**: Access

**First VLC version**: 0.5.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: FTP input

**Shortcut(s)**: -

Module: ftp

**Type**: Access output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: FTP upload output

**Shortcut(s)**: -

#### Options

-   **ftp-caching \** : Caching in ms
-   **ftp-user \** : Username *default value: "anonymous"*
-   **ftp-password \** : Password *default value: "anonymous@example.com"*
-   **ftp-account \** : Account *default value: "anonymous"*

#### Source code

-   [modules/access/ftp.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/ftp.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### galaktos {#modules-galaktos}

Module: galaktos

**Type**: Visualization

**First VLC version**: 0.8.0

**Last VLC version**: 1.0.6

**Operating system(s)**: Linux

**Description**: Galaktos visualization plugin (MilkDrop compatible)

**Shortcut(s)**: -

**This page is obsolete and kept only for historical interest.** It may document features that are obsolete, superseded, or irrelevant. Do not rely on the information here being up-to-date.

Between VLC version 1.0.6 and 1.1.0 support for this module was dropped and users were directed to projectm instead.

### gather {#modules-gather}

Module: gather

**Type**: Stream output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Recycle video and audio elementary streams when possible

**Shortcut(s)**: -

Makes it possible to stream a playlist without any noticeable interruption on input change on the client side.
The audio and video streams must all have the same characteristics (codecs, bit rate, dimensions, etc.).

#### Example

    % vlc playlist.m3u --sout "#gather:std{access=http,mux=asfh,dst=:8080}" --sout-keep

If your playlist items use different codecs or have different sizes, it is advised to transcode. For example:

    % vlc playlist.m3u --sout "#transcode{vcodec=DIV3,vb=512,width=640,height=480,acodec=mp3,ab=128,samplerate=44100,channels=2}:gather:std{access=http,mux=asfh,dst=:8080}" --sout-keep

It is unclear whether using 0 automatically sets gather automatically or not 1

See also VLC HowTo/Merge videos together

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### gaussianblur {#modules-gaussianblur}

Module: gaussianblur

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: gaussian blur video filter

**Shortcut(s)**: -

Use this filter to blur the whole video.

#### Options

-   **gaussianblur-sigma \** : The gaussian's standard deviation. *default value: 2.*

#### Example

    % vlc --video-filter "gaussianblur{sigma=3.45}" somevideo.avi

#### See also

-   Documentation:Modules/sharpen

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### glwin32 {#modules-glwin32}

Module: glwin32

**Type**: Video output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: Windows

**Description**: OpenGL video output for Windows

**Shortcut(s)**: 0, 1

#### Source code

-   [modules/video_output/win32/glwin32.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_output/win32/glwin32.c)

#### See also

-   opengl
-   glwin32
-   glx

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### glx {#modules-glx}

Module: glx

**Type**: Video output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: OpenGL X video output

**Shortcut(s)**: (none)

#### Source code

-   [modules/video_output/glx.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_output/glx.c)

#### See also

-   opengl
-   glwin32
-   glx

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### gme {#modules-gme}

Module: gme

**Type**: Access demux

**First VLC version**: 1.1.5

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Game_Music_Emu

**Shortcut(s)**: 0

Game_Music_Emu was supported before the introduction of this module in 1.1.5, according to [forum thread #78168](https://forum.videolan.org/viewtopic.php?f=2&t=78168). The option 0 would enable the module.

#### Options

None.

#### Source code

-   [modules/demux/gme.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/gme.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### goom {#modules-goom}

Module: goom

**Type**: Visualization

**First VLC version**: 0.7.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Audio visual effects filter

**Shortcut(s)**: -

This is a plugin. See 0 for an intro and 1 for development (the repo is Bazaar).

#### Source code

-   [modules/visualization/goom.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/visualization/goom.c)

### gradfun {#modules-gradfun}

Module: gradfun

**Type**: Video filter

**First VLC version**: 2.0.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Gradfun video filter

**Shortcut(s)**: -

This module is a wrapper for the gradfun filter from libav.

#### Options

-   **gradfun-radius \** : Radius in pixels *default value: 16*
-   **gradfun-strength \** : Strength used to modify the value of a pixel *default value: 1.2*

#### Source code

-   [modules/video_filter/gradfun.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/gradfun.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### gradient {#modules-gradient}

Module: gradient

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: apply gradients and edge detection filters

**Shortcut(s)**: -

Use this filter to apply gradients or do simple edge detection algorithms to the video image.

#### Options

-   **gradient-mode \** : One of "gradient", "edge" or "hough". *default value: gradient*

&nbsp;

-   **gradient-type \** : 0 to discard colors, 1 to keep colors. *default value: 0*

&nbsp;

-   **gradient-cartoon** : Apply a cartoon effect. *default value: enabled*

#### Example

    % vlc --video-filter "gradient{type=1}" somevideo.avi

#### Screenshots


**Note:** In versions prior to 0.9.0, this was part of the distort video filter.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### h26x {#modules-h26x}

Module: h26x

**Type**: Access demux

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Raw H264 and HEVC Video demuxers

**Shortcut(s)**: -

The H.265 video demuxer is a sub-module.

The H264 video demuxer has a shortcut of 0 and the HEVC/H.265 video demuxer has shortcuts of 1 and 2.

#### Options

##### H264 video demuxer

-   **h264-fps \** : Desired frame rate for the stream *default value: 0.0f*

##### HEVC/H.265 video demuxer

-   **hevc-fps \** : Desired frame rate for the stream *default value: 0.0f*

#### Source code

-   [modules/demux/mpeg/h26x.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/mpeg/h26x.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### hal {#modules-hal}

Module: hal

**Type**: Services discovery

**First VLC version**: 0.8.2

**Last VLC version**: 1.1.13

**Operating system(s)**: Linux

**Description**: [HAL](http://en.wikipedia.org/wiki/HAL_(software) "wikipedia:HAL (software)") devices detection

**Shortcut(s)**: (none)

This module was removed because it is no longer needed; it was used on \*nix systems before the advent of [udev](http://en.wikipedia.org/wiki/udev).

It was removed with this note:

    HAL is officially deprecated. The new udev discs module provide the same
    functionality in VLC. Moreover, the plugin was waking up the CPU at
    regular intervals. Last, InitDeviceValues seemed to cause problems with
    wrong disc paths being saved to vlcrc for some people

#### Options

None.

#### Source code

-   [\[9ee9e9aa659dfa5f283b28cc971eab622a7c9052\]](https://git.videolan.org/?p=vlc/vlc-0.8.git;a=blob;f=modules/services_discovery/hal.c;h=9ee9e9aa659dfa5f283b28cc971eab622a7c9052;hb=8b61d4ef6120a68ea9e4dd3865d6a35d11965e2c) (introduction)
-   [modules/services_discovery/hal.c](https://git.videolan.org/?p=vlc/vlc-0.8.git;a=blob;f=modules/services_discovery/hal.c) (vlc/vlc-0.8.git)
-   [modules/services_discovery/hal.c](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=blob;f=modules/services_discovery/hal.c) (vlc/vlc-0.9.git)
-   [modules/services_discovery/hal.c](https://git.videolan.org/?p=vlc/vlc-1.0.git;a=blob;f=modules/services_discovery/hal.c) (vlc/vlc-1.0.git)
-   [\[0565b5c2e5062b41e6e1d2b441724899bfdcf38d\]](https://git.videolan.org/?p=vlc/vlc-1.1.git;a=commitdiff;h=0565b5c2e5062b41e6e1d2b441724899bfdcf38d) (removal)

#### External links

-   [freedesktop.org - hal](https://freedesktop.org/wiki/Software/hal/)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### hotkeys {#modules-hotkeys}

Module: Hotkeys

**Type**: Control interface

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Hotkeys management interface

**Shortcut(s)**: -

This module has no shortcuts.

It's easier to set options through the GUI (see [Documentation:Hotkeys](#hotkeys)), but this module can be accessed through the core. Look for the 0{.nowrap} section after running:

    % vlc -p core --advanced --help-verbose

#### Options

None

#### Source code

-   [modules/control/hotkeys.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/control/hotkeys.c)

### http {#modules-http}

The 0 option was removed prior to VLC 2.0.0 with this commitdiff: [Unify (ACCESS\|DEMUX)\_GET_PTS_DELAY](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=31ac20b22fc37bcf78991159bf8a0f138db05b44)

The option 0 was removed from the 3.0.x and 4.0.0-dev branches with this commitdiff: [HTTP win32: use http-proxy options to setup the proxy](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=6514ed051d579972f21949be3900a48d8d62c647) with summary *Because win32/netconf is not ready*.

The option 0 changed in VLC 1.1.1 (changelog):

    libvlc_set_user_agent() configures the "user agent" strings used for some
    protocols (HTTP, PulseAudio...). This replaces the --http-user-agent and
    the former --user-agent libvlc_new() parameters.

#### HTTP

Module: http

**Type**: Access

**First VLC version**: 0.5.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: HTTP input

**Shortcut(s)**: 0, 1, 2, 3

HTTP was first supported *before* 0.5.0 (probably from the beginning).

As of VLC 0.9.0 this module accepts gzip compressed data and Digest Access Authentication.

##### Options

-   **http-reconnect** : Automatically try to reconnect in case of a sudden disconnect *default value: disabled*

#### HTTPS

Module: https

**Type**: Access

**First VLC version**: 3.0.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: HTTPS input

**Shortcut(s)**: 0

HTTPS was first supported in 0.8.1 (for http_intf). This particular sub-module was introduced in 3.0.0 for HTTP 2.0 support.

##### Options

-   **http-forward-cookies \** : Forward cookies across HTTP redirections *default value: enabled*
-   **http-referrer \** : Provide the referral URL, i.e. HTTP "[Referer](http://en.wikipedia.org/wiki/HTTP_referer)" \[*[sic](http://en.wiktionary.org/wiki/sic#Usage_notes)*\] *default value: NULL*
-   **http-user-agent \** : Override the name and version of the application as provided to the HTTP server, i.e. the HTTP "[User-Agent](http://en.wikipedia.org/wiki/User_agent)". Name and version must be separated by a forward slash, e.g. "FooBar/1.2.3" *default value: NULL*
-   **http-continuous \** : Keep reading a resource that keeps being updated (for example a JPEG file) *default value: disabled*

#### Source code

-   [modules/access/http.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/http.c) (file - HTTP module)
-   [modules/access/http/access.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/http/access.c) (file - HTTPS sub-module)
-   [modules/access/http](https://git.videolan.org/?p=vlc.git;a=tree;f=modules/access/http;hb=HEAD) (folder)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### http intf {#modules-http-intf}

Module: http

**Type**: Control interface

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Allows control of VLC over a http connection

**Shortcut(s)**: -

The **http** lua module makes it possible to Control VLC via a browser interface which can be enabled by going to Settings \> Add Interface \> Web Interface.

#### How to use

To use this interface as the primary interface, launch VLC with the parameter "-I http" or set http to be the primary interface via the preferences (see below for instructions). To launch it as a secondary interface you should launch VLC with the parameter "--extraintf=http" or set http as an extra-interface in the Preferences area mentioned above.

Now, when you start VLC, a web interface will be created and running on your computer on port 8080 (by default, but you can change this yourself). For your information, you can connect to a web server listening on an arbitrary port using 0 syntax, so you can test the VLC web interface using this URL: 1.

If you get a *401 Unauthorized* error message and you have set a password on the interface as described in #Access_control below, leave the username field blank and enter the password that you have set.

VLC 2.0 and below: If you try to access the web interface from another computer or by using your computer's IP address and you get a *403 Access Denied* error message, you have to allow access for the IP or IP range first: see #Access_control below.

The following options can be used to specify an IP and a different port on which you want to run the interface.

##### VLC 2.0.0 and later

     --http-host host
     --http-port port

or on Windows platforms:

     --http-host=host
     --http-port=port

To enable the HTTP control interface as a primary or extra interface, go to Tools → Preferences (select "All" radio-button) → Interface → Main interfaces → check "Web":

##### VLC before 2.0.0

     --http-host host:port

or on Windows platforms:

     --http-host=host:port

To enable the http interface as a primary or extra interface, go to Tools \> Preferences \> Interface \> General \> Interface module: http remote control interface. In later versions it might be Tools → Preferences (select All radiobutton) → All radio button → Interface → Main interfaces → check HTTP remote control)

#### Configure

##### Access control

###### VLC 2.1.0 and later

Access control has been simplified in VLC 2.1.0. You can restrict access to the web interface by using a simple password that can be set under Tools → Preferences (all) → Interfaces → Main interfaces → Lua → Lua HTTP → Password.

It can also be set from the command line as the option 'http-password', like so:

    --http-password

When logging in, **leave the username field blank**.

###### VLC before 2.1.0

Access control for specific IP addresses or ranges of IP addresses to the http interface can be done globally by editing "/usr/share/vlc/lua/http/.hosts" in Linux, "%PROGRAMFILES%\\VideoLAN\\VLC\\lua\\http\\.hosts" on Windows and "/Applications/VLC.app/Contents/MacOS/share/lua/http/.hosts" on Mac OS X.

The existing .hosts file contains examples and can easily customized to meet your needs. On Windows, note that you might need administrator rights to edit this file.

    Note that the global file gets overwritten when/if you reinstall/upgrade VLC.
    This is solved by some Linux distributions by symlinking the file to /etc.
    If your distribution does not do this; execute the following as root:
    mkdir /etc/vlc && cd /usr/share/vlc/lua/http/ && mv .hosts /etc/vlc && ln -s /etc/vlc/.hosts .hosts

##### Customization

It is now also possible to customize the Web interface. See the html pages in share/html (within the VLC directory for Windows, within the VLC.app package on Mac OS X and somewhere in /usr/local for Linux). This can be a very cool interface if you spent some time developing nice UI elements. If you would like to contribute a new 'Default' html interface, you are also very welcome (keep it small).

An additional theme has been created (by Lucas Steigmeyer a.k.a. Plezops) specifically for PDA's or PSP's in mind. This additional theme has a grey theme and will fit nicely on most portable versions of web browsers. The theme has the layout reordered to fit the screen better. You may download this theme from [OrrentDesign.com](http://www.orrentdesign.com/outsideResources/VLC_Graphite.zip). A readme file is included. View for install instructions and other info.

#### Notes

-   On versions of VLC (windows) that are greater than .8 (possibly previous versions, though not confirmed) the HTTP interface index file is stored in the "http" folder in the VLC folder. There is a README file that serves as documentation, VLC HTTP requests.
-   A new http interface is available since 0.8.5. However this new interface does not work on handheld PDA's running the Windows Mobile OS, it also does not work with JavaScript turned off in your browser. This old interface was available at http://\:\/old/ for releases prior to 1.0.0.
-   Since 2.0.0, the HTTP interface has been rewritten from the ground up as a lua plugin with AJAX, and the oldhttp interface no longer exists.
-   Since 2.1.0, the HTTP interface no longer uses the hosts file, but instead a password.

For more information about the HTTP interface, see the document "VLC Play-Howto", the paragraph "The HTTP interface" in chapter 4 ("Advanced use of VLC") and Documentation:Play HowTo/Building Pages for the HTTP Interface. ("See also old/outdated appendix B").

#### See also

-   Documentation:Play HowTo/Building Pages for the HTTP Interface (may be obsolete)
-   Interfaces

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### image {#modules-image}

#### Video output

Module: image

**Type**: Video output

**First VLC version**: 0.8.2

**Last VLC version**: 0.9.10

**Operating system(s)**: all

**Description**: Outputs the video images to files

**Shortcut(s)**: -

In VLC 1.0.0 the image video output was rewritten into a video-filter named scene, and the old image video output was removed.

Trivia: [the help text](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=blob;f=modules/video_output/image.c#l56) was never changed after [this commitdiff](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=5183b07c7c04e302af409fca4804e66777a6a040) changed the default values of unsigned integers 0 and 1 from 2 to 3—there was little point in fixing the help text for a deprecated module in software not yet publicly released! The coding error is absent from the current module, scene.

Option aliases 0 for 1 and 2 for 3 were deprecated in 0.9.0.

##### Options

-   **image-out-format \ {png,jpeg}** : Format of the output images *default value: png*
-   **image-out-width \** : You can enforce the image width. By default VLC will adapt to the video characteristics *default value: 0*
-   **image-out-height \** : You can enforce the image height. By default VLC will adapt to the video characteristics *default value: 0*
-   **image-out-ratio \** : Ratio of images to record. *3* means that one image out of three is recorded *default value: 3*
-   **image-out-prefix \** : Prefix of the output images filenames. Output filenames will have the "prefixNUMBER.format" form. Starting with VLC 0.9.0 you can also use format time and meta variables *default value: img*
-   **image-out-replace \** : Always write to the same file instead of creating one file per image. In this case, the number is not appended to the filename *default value: disabled*

#### Demux

Module: image

**Type**: Access demux

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Image demuxer

**Shortcut(s)**: -

##### Options

-   **image-id \** : Set the ID of the elementary stream *default value: -1*
-   **image-group \** : Set the group of the elementary stream *default value: 0*
-   **image-decode \** : Decode at the demuxer stage *default value: enabled*
-   **image-chroma \** : If non empty and 0{.variable} is true, the image will be converted to the specified chroma *default value: ""*
-   **image-duration \** : Duration in seconds before simulating an end of file. A negative value means an unlimited play time *default value: 10*
-   **image-fps \** : Frame rate of the elementary stream produced *default value: 10/1*
-   **image-realtime \** : Use real-time mode suitable for being used as a master input and real-time input slaves *default value: disabled*

#### Source code

-   [modules/video_output/image.c](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=blob;f=modules/video_output/image.c) (vlc/vlc-0.9.git) (video output)
-   [modules/demux/image.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/image.c) (image demuxer)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### invert {#modules-invert}

Module: invert

**Type**: Video filter

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: color inversion

**Shortcut(s)**: -

The invert filter inverts colors in an image.

#### Example

**VLC 0.9.0 and above**:

    $ vlc --video-filter invert somevideo.avi

**Note:** In versions prior to 0.9.0, invert was a video output filter.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### jack {#modules-jack}

This module allows VLC media player to connect to JACK Audio Connection Kit.

#### Access

Module: jack

**Type**: Access

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: Unix, Linux, BSD

**Description**: JACK input

**Shortcut(s)**: 0

The option 0 no longer exists, removed with a commitdiff entitled [Unify (ACCESS\|DEMUX)\_GET_PTS_DELAY](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=31ac20b22fc37bcf78991159bf8a0f138db05b44).

-   **jack-input-use-vlc-pace \** : Read the audio stream at VLC pace rather than Jack pace *default value: disabled*
-   **jack-input-auto-connect \** : Automatically connect VLC input ports to available output ports *default value: disabled*

#### Audio output

Module: jack

**Type**: Audio output

**First VLC version**: 0.8.5

**Last VLC version**: -

**Operating system(s)**: Unix, Linux, BSD

**Description**: JACK audio output

**Shortcut(s)**: (none)

-   **jack-auto-connect \** : If enabled, this option will automatically connect sound output to the first writable JACK clients found *default value: enabled*
-   **jack-connect-regex \** : If automatic connection is enabled, only JACK clients whose names match this [regular expression](http://en.wikipedia.org/wiki/regular_expression) will be considered for connection *default value: "system"*
-   **jack-name \** : JACK client name *default value: ""*

#### Source code

-   [modules/access/jack.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/jack.c)
-   [modules/audio_output/jack.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/audio_output/jack.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### jpeg {#modules-jpeg}

#### Decoder

Module: jpeg

**Type**: Access demux

**First VLC version**: 2.2.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: JPEG image decoder through libjpeg

**Shortcut(s)**: (none)

VLC media player previously decoded JPEGs through the images demuxer (introduced in VLC 2.0.0).

##### Options

None.

#### Encoder

Module: jpeg

**Type**: Muxer

**First VLC version**: 2.2.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: JPEG image encoder through libjpeg

**Shortcut(s)**: 0

##### Options

-   **sout-jpeg-quality \** : Quality level for encoding (this can enlarge or reduce output image size)

#### Source code

-   [modules/codec/jpeg.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/jpeg.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### kate {#modules-kate}

Module: kate

**Type**: Access demux

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Kate overlay decoder

**Shortcut(s)**: 0

The help text for this module:

    Kate is a codec for text and image based overlays.
    The Tiger rendering library is needed to render complex Kate streams, but VLC can still render static text and image based subtitles if it is not available.
    Note that changing settings below will not take effect until a new stream is played. This will hopefully be fixed soon.

#### Options

##### Basic

-   **kate-formatted \** : Kate streams allow for text formatting. VLC partly implements this, but you can choose to disable all formatting. Note that this has no effect if rendering via Tiger is enabled *default value: enabled*

##### Tiger

-   **kate-use-tiger \** : Kate streams can be rendered using the Tiger library. Disabling this will only render static text and bitmap based streams *default value: enabled*
-   **kate-tiger-quality \** : Select rendering quality, at the expense of speed. 0 is fastest, 1 is highest quality *default value: 1.0*

###### Tiger rendering defaults

-   **kate-tiger-default-font-desc \** : Which font description to use if the Kate stream does not specify particular font parameters (name, size, etc) to use. A blank name (default) will let Tiger choose font parameters where appropriate
-   **kate-tiger-default-font-effect \** : Add a font effect to text to improve readability against different backgrounds *default value: 0*
-   **kate-tiger-default-font-effect-strength \** : How pronounced to make the chosen font effect (effect dependent) *default value: 0.5*
-   **kate-tiger-default-font-color \** : Default font color to use if the Kate stream does not specify a particular font color to use *default value: 0x00ffffff (white)*
-   **kate-tiger-default-font-alpha \** : Transparency of the default font color if the Kate stream does not specify a particular font color to use (0x00 is fully transparent, 0xff is fully opaque) *default value: 0xff*
-   **kate-tiger-default-background-color \** : Default background color if the Kate stream does not specify a background color to use *default value: 0x00ffffff (white)*
-   **kate-tiger-default-background-alpha \** : Transparency of the default background color if the Kate stream does not specify a particular background color to use (0x00 is fully transparent, 0xff is fully opaque) *default value: 0x00*

#### Source code

-   [modules/codec/kate.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/kate.c)

### lirc {#modules-lirc}

Module: lirc

**Type**: Interface

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: Any that supports [lirc](http://www.lirc.org/) library

**Description**: Infrared remote control interface

**Shortcut(s)**: -

This module lets you control VLC using your infrared remote control using lirc.

#### Options

-   **lirc-file \** : Tell lirc to read this configuration file. By default it searches in the users home directory.

### live {#modules-live}

Module: live555

**Type**: Access demux

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: RTP/RTSP/SDP demuxer (using Live555)

**Shortcut(s)**: 0, 1

The 0 option was removed prior to VLC 2.0.0 with this commitdiff: [Unify (ACCESS\|DEMUX)\_GET_PTS_DELAY](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=31ac20b22fc37bcf78991159bf8a0f138db05b44)

#### Options

None.

#### Submodule

Module: live555

**Type**: Access

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: RTSP/RTP access and demux

**Shortcut(s)**: 0, 1, 2, 3

##### Options

-   **rtsp-tcp \** : Use RTP over RTSP (TCP) *default value: disabled*
-   **rtp-client-port \** : Port to use for the RTP source of the session *default value: -1*
-   **rtsp-mcast \** : Force multicast RTP via RTSP *default value: disabled*
-   **rtsp-http \** : Tunnel RTSP and RTP over HTTP *default value: disabled*
-   **rtsp-http-port \** : Port to use for tunneling the RTSP/RTP over HTTP *default value: 80*
-   **rtsp-kasenna \** : Kasenna servers use an old and nonstandard dialect of RTSP. With this parameter VLC will try this dialect, but then it cannot connect to normal RTSP servers *default value: disabled*
-   **rtsp-wmserver \** : WMServer uses a nonstandard dialect of RTSP. Selecting this parameter will tell VLC to assume some options contrary to [RFC 2326](https://tools.ietf.org/html/rfc2326) guidelines *default value: disabled*
-   **rtsp-user \** : Sets the username for the connection, if no username or password are set in the url *default value: NULL*
-   **rtsp-pwd \** : Sets the password for the connection, if no username or password are set in the url *default value: NULL*
-   **rtsp-frame-buffer-size \** : RTSP start frame buffer size of the video track, can be increased in case of broken pictures due to too small buffer *default value: 250000*

#### Source code

-   [modules/access/live555.cpp](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/live555.cpp)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### live555 {#modules-live555}

Module: live555

**Type**: Access demux

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: RTP/RTSP/SDP demuxer (using Live555)

**Shortcut(s)**: 0, 1

The 0 option was removed prior to VLC 2.0.0 with this commitdiff: [Unify (ACCESS\|DEMUX)\_GET_PTS_DELAY](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=31ac20b22fc37bcf78991159bf8a0f138db05b44)

#### Options

None.

#### Submodule

Module: live555

**Type**: Access

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: RTSP/RTP access and demux

**Shortcut(s)**: 0, 1, 2, 3

##### Options

-   **rtsp-tcp \** : Use RTP over RTSP (TCP) *default value: disabled*
-   **rtp-client-port \** : Port to use for the RTP source of the session *default value: -1*
-   **rtsp-mcast \** : Force multicast RTP via RTSP *default value: disabled*
-   **rtsp-http \** : Tunnel RTSP and RTP over HTTP *default value: disabled*
-   **rtsp-http-port \** : Port to use for tunneling the RTSP/RTP over HTTP *default value: 80*
-   **rtsp-kasenna \** : Kasenna servers use an old and nonstandard dialect of RTSP. With this parameter VLC will try this dialect, but then it cannot connect to normal RTSP servers *default value: disabled*
-   **rtsp-wmserver \** : WMServer uses a nonstandard dialect of RTSP. Selecting this parameter will tell VLC to assume some options contrary to [RFC 2326](https://tools.ietf.org/html/rfc2326) guidelines *default value: disabled*
-   **rtsp-user \** : Sets the username for the connection, if no username or password are set in the url *default value: NULL*
-   **rtsp-pwd \** : Sets the password for the connection, if no username or password are set in the url *default value: NULL*
-   **rtsp-frame-buffer-size \** : RTSP start frame buffer size of the video track, can be increased in case of broken pictures due to too small buffer *default value: 250000*

#### Source code

-   [modules/access/live555.cpp](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/live555.cpp)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### logo {#modules-logo}

See also: Documentation:Modules/erase and VLC HowTo/Add a logo

The logo filter can be used to add a logo on the video. This logo can be a static image or series of images which will be displayed alternatively. When used as a video output filter, you can move the logo with the mouse.

#### Video sub-filter

Module: logo

**Type**: Video sub-filter

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Logo sub source

**Shortcut(s)**: 0

-   **logo-file \** : Image to display. The full format is 0.
-   **logo-x \** : X offset from upper left corner. *default value: 0*
-   **logo-y \** : Y offset from upper left corner. *default value: 0*
-   **logo-position \ { 0, 1, 2, 4, 8, 5, 6, 9, 10 }** : Logo position. *default value: 5*
-   **logo-opacity \** : Logo opacity. 0 is transparent, 255 is fully opaque. *default value: 255*
-   **logo-delay \** : Global delay in [ms](http://en.wiktionary.org/wiki/ms#Translingual). Sets the duration each image will be displayed for in a loop iteration unless specified otherwise in the 0 option. *default value: 1000*
-   **logo-repeat \** : Number of loops for the logo animation. -1 for continuous, 0 to disable. *default value: -1*

#### Video output filter

Module: logo

**Type**: Video output filter

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Logo video filter

**Shortcut(s)**: 0

#### Examples

    % vlc --video-filter "logo{file=cone.png,opacity=128}" somevideo.avi

This command will display image cone.png in the video's upper right corner with 50% transparency.

    % vlc --video-filter "logo{file='cone1.png,2000,128;cone2.png,3000'}" somevideo.avi

This command will display image cone1.png for 2 seconds with 50% transparency followed by image cone2.png for 3 seconds at default transparency and loop.

#### Source code

-   [modules/spu/logo.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/spu/logo.c)

#### Appendix

-   \^ --logo-position

&nbsp;

-   **Integer**: Alignment Comment
-   **0**: Center
-   **1**: Left
-   **2**: Right
-   **4**: Top
-   **8**: Bottom
-   **5**: Top-Left 4 + 1
-   **6**: Top-Right 4 + 2
-   **9**: Bottom-Left 8 + 1
-   **10**: Bottom-Right 8 + 2
-   **3**: n/a contradictory
-   **7**: n/a contradictory

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### lua {#modules-lua}

Module: lua

**Type**: Interface

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: [Lua](http://en.wikipedia.org/wiki/Lua_(programming_language) "wikipedia:Lua (programming language)") interpreter

**Shortcut(s)**: 0

#### Options

-   **lua-intf \** : [Lua](http://en.wikipedia.org/wiki/Lua_(programming_language) "wikipedia:Lua (programming language)") interface module to load *default value: "dummy"*
-   **lua-config \** : Lua interface configuration string. Format is: 0. *default value: ""*

#### Lua HTTP

Module: luahttp

**Type**: Interface

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Lua HTTP

**Shortcut(s)**: 0, 1

-   **http-password \** : A single password restricts access to this interface. *default value: NULL*
-   **http-src \** : Source directory *default value: NULL*
-   **http-index \** : Allow to build directory index *default value: disabled*

#### Lua Telnet

Module: luatelnet

**Type**: Interface

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Lua Telnet

**Shortcut(s)**: 0, 1

-   **telnet-host \** : This is the host on which the interface will listen. It defaults to all network interfaces (0.0.0.0). If you want this interface to be available only on the local machine, enter "[127.0.0.1](http://en.wikipedia.org/wiki/localhost)". *default value: "localhost"*
-   **telnet-port \** : This is the TCP port on which this interface will listen. It defaults to 4212. *default value: 4212*
-   **telnet-password \** : A single password restricts access to this interface. *default value: NULL*

#### Lua SD Module

Module: luasd

**Type**: Services discovery

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Lua SD Module

**Shortcut(s)**: 0

-   **lua-sd \** : *default value: ""*

#### Other submodules

-   **Name**: Description Capability Shortcut
-   **Lua Meta Fetcher**: Fetch meta data using lua scripts meta fetcher (none)
-   **Lua Meta Reader**: Read meta data using lua scripts meta reader (none)
-   **Lua Playlist**: Lua Playlist Parser Interface stream_filter luaplaylist
-   **Lua Art**: Fetch artwork using lua scripts art finder (none)
-   **Lua Extension**: Lua Extension extension luaextension

#### Source code

-   [modules/lua/vlc.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/lua/vlc.c)
-   [modules/lua](https://git.videolan.org/?p=vlc.git;a=tree;f=modules/lua;hb=HEAD)

#### See also

-   Documentation:Building Lua Playlist Scripts
-   Interfaces
-   ncurses

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### macos {#modules-macos}

**User GuideInstallationVLC**

-   About
-   Check ...
-   Preferences
-   Add Interface

**File**

-   Open File
-   Quick Open File
-   Open Disk
-   Open Network
-   Services Discovery
-   Streaming & Export Wizard
-   Save Playlist

**Edit**

-   Cut
-   Copy
-   Paste
-   Clear
-   Select All
-   Special Characters...

**Playback**

-   Play Stop
-   Step Forward
-   Jump to time
-   Previous
-   Random
-   Repeat
-   Add Folder
-   Program
-   Title
-   Chapter

**Audio**

-   Volume Up/Down/Mute
-   Visualisations
-   Audio Device
-   Audio Track
-   Audio Channels

**Video**

-   Half/Normal/Double/Fit to screen
-   Full Screen
-   Float on top
-   Snapshot
-   Deinterlace
-   Post-processing
-   Aspect Ratio
-   Crop
-   Video device
-   Video track
-   Subtitles track

**Window**

-   Minimise
-   Close
-   Controller
-   Playlist
-   Information
-   Extended Controls
-   Equaliser
-   Bookmarks
-   Messages
-   Bring All to Front

**Help**

-   Read Me
-   Online Documentation
-   Report a bug
-   Online Forum
-   Make a donation
-   VideoLAN licence
-   Licence

**Frequently Asked Questions**

 view this alone

#### Graphical Interface

Many people who want to use VLC media player on macOS will be intending to use the standard graphical interface that is provided by VLC. The standard interface consists of the eight menus in the menu bar and the 'VLC - Controller' window that opens up by default. This section outlines what VLC can do for you (at V0.8.6a current active is V3.0.12) and will be completed as I check the use of menu options.

The ten menu bar options are listed below along with the main interesting capabilities under each menu item:

-   VLC which allows you to check for an updated application, to access the preferences, and to add an interface.
-   File which allows you to open a media file, or an associated file (such as subtitles). It also has a wizard to allow the streaming of video, or the capturing of a streamed video to a file.
-   Edit which does nothing VLC-specific.
-   View which allows you to hide or show various options like previous/next buttons, shuffle and repeat, audio effects, sidebar, as well as customize what you see in 'playlist table columns'.
-   Playback allows you do do all the things you might expect from a video player; some of these features are duplicated graphically in the 'Controller' window.
-   Audio allows you to control the audio level, as well as the output device and the audio track to use from the input.
-   Video allows you to control the video display on your screen, as well as which device to display on, and which video source to show in that display.
-   Subtitles allows you to add subtitle files to your video, as well as change the appearance of subtitle text for your video.
-   Window allows you to display seven helper windows that will display information about VLC's activity, and control more detail of that activity.
-   Help gives access to the help that came with the installation, the help info on the VideoLAN site, and access to interaction mechanisms with the VLC developers.

In general, many users find that they can get what they want from VLC "straight out of the box", and may only want more advanced controls after becoming familiar with the regular interface.

#### Keyboard Shortcuts

You can find most of the keyboard shortcuts by taking a look at the menus. Additional hotkeys are defined in the section "Hotkeys" of your VLC preferences.

Some handy key combos are:

-   Spacebar – pause/unpause the video
-   ⌘ + F – toggle fullscreen (Escape will also exit fullscreen)
-   ⌘ + Shift + left/right arrow keys – jump the video back/forward about a minute
-   ⌘ + Ctrl + left/right arrow keys – jump the video back/forward about ten seconds
-   When watching a DVD, and the video window is the front-most window, arrow keys and the enter key will allow you to navigate the DVD menus
-   F key – Decrease Audio Delay in milliseconds
-   G key – Increase Audio Delay in milliseconds
-   H key – Decrease Subtitle Delay
-   J key – Increase Subtitle Delay

#### Latest developments

##### Streaming Wizard

A streaming wizard has been available since the VLC media player 0.8.4 release. This is available under the "File" menu.

##### Command line

You can run VLC on macOS using a terminal application (for example Terminal.app in /Applications/Utilities) with the following command:

    $ /Applications/VLC.app/Contents/MacOS/VLC [your options, "--intf=rc" for example]

On most Bourne-like shells, you can set an alias to just vlc with the following command:

    $ alias vlc='/Applications/VLC.app/Contents/MacOS/VLC'

It can be helpful to add this command to your shell setup file.

This option can also be activated from the "VLC" menu.

###### Command line examples

\~ will expand to 0{.sample}

Following command does this: Transform video-filter (flip vertically), transcode (save) to file.

    $ /Applications/VLC.app/Contents/MacOS/VLC -I rc --vout-filter=transform --transform-type=vflip /Movie.mov --sout='#transcode{vcodec=h264,vb=800,scale=1,acodec=mp4a,ab=128,channels=2,samplerate=44100}:std{access=file,mux=ts,dst=/output.mp4}

-I rc is so that it doesn't open the GUI, but stays on the command line version --vout-filter defines the filter to use --transform-type defines the attributes of the transform filter /Movie.mov is the file to convert --sout= is the stream output chain /output.mp4 is the output file name

###### Another Example

I had a heck of a time getting this to work the way I wanted it. I kept attempting a command-line execution of VLC to only get the following response (not what I wanted):

    VLC media player 2.0.2 Twoflower (revision 2.0.2-9-gd1b4a63)
    [0x100283cd0] [cli] lua interface: Listening on host "*console".
    VLC media player 2.0.2 Twoflower
    Command Line Interface initialized. Type `help' for help.

What I wasn't doing apparently was specifying the location of the source movie.

Eventually I ran this:

    $ vlc ~/Desktop/my_movie.mp4 --intf=rc --sout "#transcode{vcodec=VP80,vb=800,scale=1,acodec=vorbis,ab=128,channels=2}:std{access=file,mux="ffmpeg{mux=webm}",dst=my_first_transcoded_movie.webm}"

**HINT**:
This would be the same as if you didn't have an alias for vlc that pointed to the actual Applications executable:

    $ /Applications/VLC.app/Contents/MacOS/VLC ~/Desktop/mymovie.mp4 --intf=rc --sout "#transcode{vcodec=VP80,vb=800,scale=1,acodec=vorbis,ab=128,channels=2}:std{access=file,mux="ffmpeg{mux=webm}",dst=my_first_transcoded_movie.webm}"

Hopefully, I'll add to this post when the transcoding finishes and I see my results (I have no idea if I've got the correct options for vp8/vorbis webm-container transcoding).....

##### No Dock

In previous versions you can replace the 0 at the end of the path with 1 to suppress the launch of any Mac-like interface (VLC wouldn't even appear in the Dock then) or if transcoding from the command-line crashed with a 2{.sample}.

**This does not work anymore** (see Forum thread #58378)

As given by Command-line interface#macOS, specify the option 0 followed by the interface you want to add e.g. 1.

#### Need Help?

See the FAQ on macOS only issues or the Common Problems pages.

### macosx gui {#modules-macosx-gui}

**User GuideInstallationVLC**

-   About
-   Check ...
-   Preferences
-   Add Interface

**File**

-   Open File
-   Quick Open File
-   Open Disk
-   Open Network
-   Services Discovery
-   Streaming & Export Wizard
-   Save Playlist

**Edit**

-   Cut
-   Copy
-   Paste
-   Clear
-   Select All
-   Special Characters...

**Playback**

-   Play Stop
-   Step Forward
-   Jump to time
-   Previous
-   Random
-   Repeat
-   Add Folder
-   Program
-   Title
-   Chapter

**Audio**

-   Volume Up/Down/Mute
-   Visualisations
-   Audio Device
-   Audio Track
-   Audio Channels

**Video**

-   Half/Normal/Double/Fit to screen
-   Full Screen
-   Float on top
-   Snapshot
-   Deinterlace
-   Post-processing
-   Aspect Ratio
-   Crop
-   Video device
-   Video track
-   Subtitles track

**Window**

-   Minimise
-   Close
-   Controller
-   Playlist
-   Information
-   Extended Controls
-   Equaliser
-   Bookmarks
-   Messages
-   Bring All to Front

**Help**

-   Read Me
-   Online Documentation
-   Report a bug
-   Online Forum
-   Make a donation
-   VideoLAN licence
-   Licence

**Frequently Asked Questions**

 view this alone

#### Graphical Interface

Many people who want to use VLC media player on macOS will be intending to use the standard graphical interface that is provided by VLC. The standard interface consists of the eight menus in the menu bar and the 'VLC - Controller' window that opens up by default. This section outlines what VLC can do for you (at V0.8.6a current active is V3.0.12) and will be completed as I check the use of menu options.

The ten menu bar options are listed below along with the main interesting capabilities under each menu item:

-   VLC which allows you to check for an updated application, to access the preferences, and to add an interface.
-   File which allows you to open a media file, or an associated file (such as subtitles). It also has a wizard to allow the streaming of video, or the capturing of a streamed video to a file.
-   Edit which does nothing VLC-specific.
-   View which allows you to hide or show various options like previous/next buttons, shuffle and repeat, audio effects, sidebar, as well as customize what you see in 'playlist table columns'.
-   Playback allows you do do all the things you might expect from a video player; some of these features are duplicated graphically in the 'Controller' window.
-   Audio allows you to control the audio level, as well as the output device and the audio track to use from the input.
-   Video allows you to control the video display on your screen, as well as which device to display on, and which video source to show in that display.
-   Subtitles allows you to add subtitle files to your video, as well as change the appearance of subtitle text for your video.
-   Window allows you to display seven helper windows that will display information about VLC's activity, and control more detail of that activity.
-   Help gives access to the help that came with the installation, the help info on the VideoLAN site, and access to interaction mechanisms with the VLC developers.

In general, many users find that they can get what they want from VLC "straight out of the box", and may only want more advanced controls after becoming familiar with the regular interface.

#### Keyboard Shortcuts

You can find most of the keyboard shortcuts by taking a look at the menus. Additional hotkeys are defined in the section "Hotkeys" of your VLC preferences.

Some handy key combos are:

-   Spacebar – pause/unpause the video
-   ⌘ + F – toggle fullscreen (Escape will also exit fullscreen)
-   ⌘ + Shift + left/right arrow keys – jump the video back/forward about a minute
-   ⌘ + Ctrl + left/right arrow keys – jump the video back/forward about ten seconds
-   When watching a DVD, and the video window is the front-most window, arrow keys and the enter key will allow you to navigate the DVD menus
-   F key – Decrease Audio Delay in milliseconds
-   G key – Increase Audio Delay in milliseconds
-   H key – Decrease Subtitle Delay
-   J key – Increase Subtitle Delay

#### Latest developments

##### Streaming Wizard

A streaming wizard has been available since the VLC media player 0.8.4 release. This is available under the "File" menu.

##### Command line

You can run VLC on macOS using a terminal application (for example Terminal.app in /Applications/Utilities) with the following command:

    $ /Applications/VLC.app/Contents/MacOS/VLC [your options, "--intf=rc" for example]

On most Bourne-like shells, you can set an alias to just vlc with the following command:

    $ alias vlc='/Applications/VLC.app/Contents/MacOS/VLC'

It can be helpful to add this command to your shell setup file.

This option can also be activated from the "VLC" menu.

###### Command line examples

\~ will expand to 0{.sample}

Following command does this: Transform video-filter (flip vertically), transcode (save) to file.

    $ /Applications/VLC.app/Contents/MacOS/VLC -I rc --vout-filter=transform --transform-type=vflip /Movie.mov --sout='#transcode{vcodec=h264,vb=800,scale=1,acodec=mp4a,ab=128,channels=2,samplerate=44100}:std{access=file,mux=ts,dst=/output.mp4}

-I rc is so that it doesn't open the GUI, but stays on the command line version --vout-filter defines the filter to use --transform-type defines the attributes of the transform filter /Movie.mov is the file to convert --sout= is the stream output chain /output.mp4 is the output file name

###### Another Example

I had a heck of a time getting this to work the way I wanted it. I kept attempting a command-line execution of VLC to only get the following response (not what I wanted):

    VLC media player 2.0.2 Twoflower (revision 2.0.2-9-gd1b4a63)
    [0x100283cd0] [cli] lua interface: Listening on host "*console".
    VLC media player 2.0.2 Twoflower
    Command Line Interface initialized. Type `help' for help.

What I wasn't doing apparently was specifying the location of the source movie.

Eventually I ran this:

    $ vlc ~/Desktop/my_movie.mp4 --intf=rc --sout "#transcode{vcodec=VP80,vb=800,scale=1,acodec=vorbis,ab=128,channels=2}:std{access=file,mux="ffmpeg{mux=webm}",dst=my_first_transcoded_movie.webm}"

**HINT**:
This would be the same as if you didn't have an alias for vlc that pointed to the actual Applications executable:

    $ /Applications/VLC.app/Contents/MacOS/VLC ~/Desktop/mymovie.mp4 --intf=rc --sout "#transcode{vcodec=VP80,vb=800,scale=1,acodec=vorbis,ab=128,channels=2}:std{access=file,mux="ffmpeg{mux=webm}",dst=my_first_transcoded_movie.webm}"

Hopefully, I'll add to this post when the transcoding finishes and I see my results (I have no idea if I've got the correct options for vp8/vorbis webm-container transcoding).....

##### No Dock

In previous versions you can replace the 0 at the end of the path with 1 to suppress the launch of any Mac-like interface (VLC wouldn't even appear in the Dock then) or if transcoding from the command-line crashed with a 2{.sample}.

**This does not work anymore** (see Forum thread #58378)

As given by Command-line interface#macOS, specify the option 0 followed by the interface you want to add e.g. 1.

#### Need Help?

See the FAQ on macOS only issues or the Common Problems pages.

### magnify {#modules-magnify}

Module: magnify

**Type**: Video output filter

**First VLC version**: 0.8.5

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Interactive magnification filter

**Shortcut(s)**: -

You can use this module to zoom on parts of a video. It is controlled using buttons drawn directly on the video output.

#### Screenshot

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### marq {#modules-marq}

Module: marq

**Type**: Video sub-filter

**First VLC version**: 0.8.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Overlays text on the video

**Shortcut(s)**: 0, 1

The *marq* subtitle filter can be used to display text on a video. The time filter was merged with this filter in version 0.9.0, adding time format string recognition. There are two methods of specifying position: coordinate and (since VLC 0.9.0) numbered positions.

#### Options

-   **marq-marquee \** : Marquee text to display. *default value: VLC*
-   **marq-file \** : File to read the marquee text from. *default value: NULL*

##### Position

-   **marq-x \** : X offset, from the left screen edge. *default value: 0*
-   **marq-y \** : Y offset, down from the top. *default value: 0*
-   **marq-position \** : Marquee position: 0=center, 1=left, 2=right, 4=top, 8=bottom, you can also use combinations of these values, eg 6 = top-right. *default value: -1*

##### Font

-   **marq-opacity \** : Opacity (inverse of transparency) of overlaid text. 0 = transparent, 255 = totally opaque. *default value: 255*
-   **marq-color \ { 0x000000, 0x808080, 0xC0C0C0, 0xFFFFFF, 0x800000, 0xFF0000, 0xFF00FF, 0xFFFF00, 0x808000, 0x008000, 0x008080, 0x00FF00, 0x800080, 0x000080, 0x0000FF, 0x00FFFF }** : Color^(**key**)^ of the text that will be rendered on the video. This must be an hexadecimal (like HTML colors). The first two chars are for red, then green, then blue. *default value: 0xFFFFFF*
-   **marq-size \** : Font size, in pixels. 0 uses the default font size. *default value: 0*

##### Misc

-   **marq-timeout \** : Number of milliseconds the marquee must remain displayed. 0 means forever. *default value: 0*
-   **marq-refresh \** : Number of milliseconds between string updates. This is mainly useful when using meta data or time format string sequences. *default value: 1000*

#### Examples

##### Versions 2.0 and later

Example command line use **(VLC 2.0.0 and newer)**:

    % vlc '--sub-source=marq{marquee="%Y-%m-%d,%H:%M:%S",position=9,color=0xFFFF00,size=12}' somevideo.avi

This example displays the current date and time in yellow in the top left corner of video.

The equivalent long form would be;

    % vlc --sub-source=marq --marq-marquee="%Y-%m-%d,%H:%M:%S" --marq-position=9 --marq-color=0xFFFFFF --marq-size=12 somevideo.avi

##### Versions 0.9.0 to 1.1.13

    $ vlc --sub-filter 'marq{marquee=$t ($P%%),color=0xFFFF00}:marq{marquee=%H:%M:%S,position=6}' somevideo.avi

This command line will show the stream's title (0) and current position (1) in the upper left corner and the current time in the upper right corner. The *single* quotes 2 enclose our 3 characters to prevent them from being interpreted as Bash variables.
On Windows the command line would be:

    > "%PROGRAMFILES%\VideoLAN\VLC\vlc.exe" --sub-filter=marq{marquee=$t ($P%%),color=0xFFFF00}:marq{marquee=%H:%m%s,position=6} somevideo.avi

#### Gallery

-

    Marq can be chained, allowing several marquees to be displayed at the same time.

#### See also

-   Documentation:Format String
-   Documentation:Modules/rss

#### Source code

-   [modules/spu/marq.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/spu/marq.c)
-   [modules/video_filter/marq.c](https://git.videolan.org/?p=vlc/vlc-0.8.git;a=blob;f=modules/video_filter/marq.c) (vlc/vlc-0.8.git)

#### Appendix

-   \^ --marq-color

&nbsp;

-   **Sample**: Colour Hex code
-   **Black**: 0
-   **Gray**: 0
-   **Silver**: 0
-   **White**: 0
-   **Maroon**: 0
-   **Red**: 0
-   **Fuchsia**: 0
-   **Yellow**: 0
-   **Olive**: 0
-   **Green**: 0
-   **Teal**: 0
-   **Lime**: 0
-   **Purple**: 0
-   **Navy**: 0
-   **Blue**: 0
-   **Aqua**: 0

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### mjpeg {#modules-mjpeg}

#### Demux

Module: mjpeg

**Type**: Access demux

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: M-JPEG camera demuxer

**Shortcut(s)**: -

##### Options

-   **mjpeg-fps \** : This is the desired frame rate when playing MJPEG from a file. Use 0 (this is the default value) for a live stream (from a camera) *default value: 0*

#### Packetizer

Module: mjpeg

**Type**: Packetizer

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: MJPEG video packetizer

**Shortcut(s)**: (none)

##### Options

None.

#### Source code

-   [modules/demux/mjpeg.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/mjpeg.c)
-   [modules/packetizer/mjpeg.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/packetizer/mjpeg.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### mkv {#modules-mkv}

Module: mkv

**Type**: Access

**First VLC version**: 0.6.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Matroska stream demuxer

**Shortcut(s)**: 0, 1

-   **mkv-use-ordered-chapters \** : Play chapters in the order specified in the segment *default value: enabled*
-   **mkv-use-chapter-codec \** : Use chapter codecs found in the segment *default value: enabled*
-   **mkv-preload-local-dir \** : Preload matroska files in the same directory to find linked segments (not good for broken files) *default value: enabled*
-   **mkv-seek-percent \** : Seek based on percent not time *default value: disabled*
-   **mkv-use-dummy \** : Read and discard unknown [EBML](http://en.wikipedia.org/wiki/Extensible_Binary_Meta_Language) elements (not good for broken files) *default value: disabled*
-   **mkv-preload-clusters \** : Find all cluster positions by jumping cluster-to-cluster before playback *default value: disabled*

#### Source code

-   [modules/demux/mkv](https://git.videolan.org/?p=vlc.git;a=tree;f=modules/demux/mkv;hb=HEAD) (folder)
-   [modules/demux/mkv/mkv.cpp](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/mkv/mkv.cpp) (main file)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### mmdevice {#modules-mmdevice}

Module: mmdevice

**Type**: Audio output

**First VLC version**: 2.0

**Last VLC version**: -

**Operating system(s)**: Windows

**Description**: Windows Multimedia Device API audio output plugin

**Shortcut(s)**: -

#### Introduction

This is the latest audio output module, starting from VLC 2.0. It uses the modern Windows Multimedia Device API introduced in Windows Vista and can offer better quality than the waveout module.

It is the only audio output permitted under Windows 8/RT.

#### Options

-   TODO

### mms {#modules-mms}

Module: mms

**Type**: Access

**First VLC version**: 0.5.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: MMS input

**Shortcut(s)**: 0, 1, 2, 3

Handles Microsoft Media Server UDP, TCP and HTTP variants, including the ability to open mms:// and mmsh:// MRLs.

In the source code for mms module it says:

     * NOTES:
     *  MMSProtocole documentation found at 0

get.to/sdp is now located at sdp.ppona.com. This document is pertinent: ([MMSprotocol.pdf](http://sdp.ppona.com/zipfiles/MMSprotocol.pdf) or [archived copy](https://archive.today/QClst))

#### Options

-   **mms-caching \** : Caching in ms
-   **mms-all** : Force selection of all streams *default value: disabled*
-   **mms-maxbitrate \** : Select the stream with the maximum bitrate under this limit *default value: 0*
-   **mmsh-proxy \** : HTTP proxy for the HTTP MMS variant. 012 *default value: ""*

#### Source code

-   [modules/access/mss](https://git.videolan.org/?p=vlc.git;a=tree;f=modules/access/mss;hb=HEAD) (folder)
-   [modules/access/mms/mms.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/mms/mms.c) (main file)
-   [modules/access/mms/mmsh.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/mms/mmsh.c) (MMS over HTTP)
-   [modules/access/mms/mmstu.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/mms/mmstu.c) (MMS over TCP or UDP)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### mod {#modules-mod}

Module: mod

**Type**: Access demux

**First VLC version**: 0.7.1

**Last VLC version**: -

**Operating system(s)**: all

**Description**: MOD demuxer (libmodplug)

**Shortcut(s)**: -

#### Options

-   **mod-noisereduction \** : Enable noise reduction algorithm. *default value: enabled*
-   **mod-reverb \** : Enable reverberation. *default value: disabled*
-   **mod-reverb-level \** : Reverberation level. *default value: 0*
-   **mod-reverb-delay \** : Reverberation delay, in [ms](http://en.wiktionary.org/wiki/ms). Usual values are from 40 to 200ms. *default value: 40*
-   **mod-megabass \** : Enable megabass mode. *default value: disabled*
-   **mod-megabass-level \** : Megabass mode level. *default value: 0*
-   **mod-megabass-range \** : Megabass mode cutoff frequency, in Hz. This is the maximum frequency for which the megabass effect applies. *default value: 10*
-   **mod-surround \** : Surround effect. *default value: disabled*
-   **mod-surround-level \** : Surround effect level. *default value: 0*
-   **mod-surround-delay \** : Surround delay, in [ms](http://en.wiktionary.org/wiki/ms). Usual values are from 5 to 40 ms. *default value: 5*

#### Source code

-   [modules/demux/mod.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/mod.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Modules {#modules}

This page lists most of the [modules](https://git.videolan.org/?p=vlc.git;a=tree;f=modules;hb=HEAD) present in the official VLC source code. Understanding these pages might require that you know about VLC and its command line usage. It is recommended that you read the other documentation first.

To list all the available modules in your VLC build, use:

    % vlc --list

To list a module's configuration options, use:

    % vlc -p  --advanced --help-verbose

#### Interfaces

##### Graphical

-   macOS
-   Qt4
-   Skins2
-   wxWidgets (up to 0.9)

##### Text

-   remote control (rc)
-   telnet (until 2.0.0)
-   ncurses

&nbsp;

-   lua

##### Other

-   HTTP
-   Hotkeys
-   lirc
-   osc
-   Wiimote

#### Outputs

##### Audio Output

-   file
-   portaudio (up to 2.0)
-   SDL (up to 1.1.13)

Android specific:

-   OpenSL ES

Linux specific:

-   Alsa
-   aRts (up to 0.9.10)
-   esound (up to 0.9.10)
-   jack
-   pulse
-   OSS

macOS specific:

-   audioqueue (up to 2.2.8)
-   auhal (HAL AudioUnit)

Windows specific:

-   DirectX
-   WASAPI
-   waveout

##### Video Output

-   ASCII Art
-   Colored ASCII Art
-   Image (up to 0.9.10)
-   OpenGL
-   SDL (up to 2.2.8)

Linux specific:

-   Direct framebuffer (up to 2.2.8)
-   Framebuffer
-   OpenGL (GLX)
-   X11
-   XVideo

Windows specific:

-   Direct3D
-   DirectX
-   OpenGL for windows
-   Windows GDI

##### Stream Output

-   Autodel
-   Delay
-   Description
-   Display
-   Dummy
-   Duplicate
-   Elementary Stream (es)
-   Gather
-   RTP
-   Standard (std)
-   Switcher (up to 2.0.9)
-   Transcode
-   Transrate (up to 1.0.2)

The following are for use in the mosaic framework only:

-   Bridge In
-   Bridge Out
-   Mosaic Bridge

#### Filters

##### Audio Filters

##### Video Filters

-   Adjust
-   Anaglyph 3D
-   AtmoLight (up to 3.0.0)
-   Color Threshold
-   Distort (up to 0.8.6 - split into various)
-   Logo Erase
-   Extract
-   Freeze
-   Gaussian Blur
-   Gradfun
-   Gradient
-   Invert
-   Motion Blur
-   Noise (up to 1.1.13)
-   Oldmovie
-   Posterize
-   Psychedelic
-   Ripple
-   Rotate
-   Scene
-   Sepia
-   Sharpen
-   VHS
-   Wave

The following video filters are for use in transcode only:

-   Canvas
-   Crop Padd

The following video filters are for use in the mosaic framework only:

-   Alpha mask
-   Blue Screen

##### Video Sub-Filters

-   Logo
-   Marq
-   Mosaic
-   RSS
-   Subsdelay
-   Time (up to 0.8.6 - merged with marq)

##### Video Output Filters

-   Crop
-   Deinterlace
-   Logo
-   Magnify
-   Puzzle
-   Transform

###### Video Splitters

-   Clone
-   Panoramix
-   Wall

##### Visualizations

-   Galaktos (up to 1.0.6)
-   Goom
-   ProjectM
-   Visual
-   Vovoid VSXu

##### Access Filters

-   Bandwidth
-   Dump
-   Record
-   Timeshift (up to 0.9.9 - moved to core)

#### Other

##### Accesses

-   CD Input
-   Directory
-   DVDnav Input - DVD with menus
-   DVDRead Input - DVD without menus
-   Fake (up to 0.9.0) - presents a static image as a video stream
-   File Input - for reading local files
-   FTP Input
-   H.264 Video
-   HTTP Input
-   jpeg
-   mjpeg
-   Matroska stream
-   MMS - for reading from the MicroSoft Media Server
-   Raw Video - streams of bitmap images
-   RTP Input
-   RTSP
-   sdp
-   Screen Input - screen feed
-   UDP Input
-   VCD

Linux specific:

-   DC1394
-   DVB Input
-   PVR (IVTV MPEG Encoding Card Input) (up to 2.0.9)
-   DV (through libdv)
-   Video4Linux (v4l) (up to 1.1.13)
-   Video4Linux2 (v4l2)

Windows specific:

-   BDA
-   DirectShow

macOS specific:

-   EyeTV (up to 2.2.8) - reads DVB streams from the proprietary EyeTV.app; requires a plugin
-   qtcapture (up to 2.2.8) - reads uncompressed video from internal iSights
-   qtsound
-   avcapture

##### Access Outputs

-   shout (shoutcast/icecast)

##### Codecs

###### Audio

-   a52
-   flac
-   mpc - Musepack
-   ogg
-   vorbis
-   wav

###### Video

-   h26x
-   nsv
-   schroedinger
-   vpx

###### Subtitles

-   kate
-   subtitle
-   telx

##### Demuxers

-   avcodec ("FFmpeg")

###### Playlist

-   playlist (formats are read with sub-modules)

##### Muxers

-   asf
-   avformat
-   avi
-   daala
-   mp4
-   mpjpeg
-   ogg
-   schroedinger
-   vpx
-   wav

##### Service Discovery

-   Bonjour
-   DAAP
-   HAL (up to 1.1.13)
-   SAP
-   Shoutcast
-   podcast
-   UPnP

##### Misc

-   Motion control
-   Netsync

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### mosaic {#modules-mosaic}

**Mosaic framework (How-To)Modules:** mosaic (mosaic-bridge • bridge-in • bridge-out) • alphamask • bluescreen

Module: mosaic

**Type**: Video sub-filter

**First VLC version**: 0.8.2

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Mosaic video sub source

**Shortcut(s)**: -

Use this filter to blend videos on top of another video. This can be used to create TV channels mosaics, setup a weather channel-like stream (with the bluescreen video filter) and lots of other fun stuff.

Since VLC 0.8.6, you can also use the HTTP interface's mosaic wizard to configure a mosaic easily.

#### Options

-   **mosaic-alpha \** : Alpha blending (transparency) of the mosaic foreground pictures. 0 means transparent, 255 opaque. *default value: 255*
-   **mosaic-height \<integer \[0 .. 0{.variable}\]\>** : Total height of the mosaic, in pixels. *default value: 100*
-   **mosaic-width \<integer \[0 .. 0{.variable}\]\>** : Total width of the mosaic, in pixels. *default value: 100*
-   **mosaic-align \ { 0, 1, 2, 4, 8, 5, 6, 9, 10 }** : You can enforce the mosaic alignment on the parent video. *default value: 5*
-   **mosaic-xoffset \<integer \[0 .. 0{.variable}\]\>** : X Coordinate of the top-left corner of the mosaic. *default value: 0*
-   **mosaic-yoffset \<integer \[0 .. 0{.variable}\]\>** : Y Coordinate of the top-left corner of the mosaic. *default value: 0*
-   **mosaic-borderw \<integer \[0 .. 0{.variable}\]\>** : Border width between mosaic elements, in pixels. *default value: 0*
-   **mosaic-borderh \<integer \[0 .. 0{.variable}\]\>** : Border height between mosaic elements, in pixels. *default value: 0*
-   **mosaic-position \ { 0, 1, 2 }** : Positioning method of the mosaic elements. Use 0 to position the elements automatically on the grid, 1 to position the elements in fixed positions on the grid (see mosaic-order) and 2 to use grid-independent offsets (see mosaic-offsets). *default value: 0*
-   **mosaic-rows \<integer \[1 .. 0{.variable}\]\>** : Number of image rows in the mosaic (only used if positioning method is set to "fixed"). *default value: 2*
-   **mosaic-cols \<integer \[1 .. 0{.variable}\]\>** : Number of image columns in the mosaic (only used if positioning method is set to "fixed"). *default value: 2*
-   **mosaic-keep-aspect-ratio \** : Keep the original aspect ratio when resizing mosaic elements. *default value: disabled*
-   **mosaic-keep-picture \** : Do not resize or do any other transformation on the mosaic pictures. Should be enabled when using the mosaic-bridge's resizing options. *default value: disabled*
-   **mosaic-order \** : You can enforce the order of the elements on the mosaic. You must give a comma-separated list of picture ID(s) (For example: tf1,fr2,fr3,m6). These IDs are assigned in the mosaic-bridge module. *default value: ""*
-   **mosaic-offsets \** : You can enforce the 0 offsets of the elements on the mosaic (only used if positioning method is set to "offsets"). You must give a comma-separated list of coordinates. For example: 10,10,150,10 if you want to position the first picture at coordinates 1 and the second one at coordinates 2. *default value: ""*
-   **mosaic-delay \** : Pictures coming from the mosaic elements will be delayed according to this value (in milliseconds). For high values you will need to raise caching at input. *default value: 0*

#### Source code

-   [modules/spu/mosaic.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/spu/mosaic.c)

#### Appendix

-   \^ --mosaic-align

&nbsp;

-   **Integer**: Alignment Comment
-   **0**: Center
-   **1**: Left
-   **2**: Right
-   **4**: Top
-   **8**: Bottom
-   **5**: Top-Left 4 + 1
-   **6**: Top-Right 4 + 2
-   **9**: Bottom-Left 8 + 1
-   **10**: Bottom-Right 8 + 2
-   **3**: n/a contradictory
-   **7**: n/a contradictory

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### mosaic / bridge {#modules-mosaic-bridge}

**Mosaic framework (How-To)Modules:** mosaic (mosaic-bridge • bridge-in • bridge-out) • alphamask • bluescreen

Module: mosaic-bridge

**Type**: Stream output

**First VLC version**: 0.8.2

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Send a picture to the mosaic framework

**Shortcut(s)**: 0

Use this filter to send a picture to the mosaic framework. Processing can be applied before sending the picture, such as resizing, chroma conversion and video filters.

#### Options

-   **sout-mosaic-bridge-id \** : Specify an identifier string for this subpicture. Used by clients of the mosaic framework to identify the picture's source. *default value: "Id"*
-   **sout-mosaic-bridge-width \** : Resize video to this width if value is non-zero. Make sure to use the **mosaic-keep-picture** option to prevent the mosaic filter from resizing a second time. *default value: 0*
-   **sout-mosaic-bridge-height \** : Resize video to this height if value is non-zero. Make sure to use the **mosaic-keep-picture** option to prevent the mosaic filter from resizing a second time. *default value: 0*
-   **sout-mosaic-bridge-sar \** : Sample aspect ratio of the destination. *default value: "1:1"*
-   **sout-mosaic-bridge-chroma \** : Force the use of a specific chroma. Use YUVA if you're planning to use the alphamask or bluescreen video filter. *default value: "I420"*
-   **sout-mosaic-bridge-vfilter \** : Video filter chain to apply after resizing and chroma conversion. *default value: ""*
-   **sout-mosaic-bridge-alpha \** : Transparency of the mosaic picture. *default value: 255*
-   **sout-mosaic-bridge-x \** : X coordinate of the upper left corner in the mosaic if non-negative. *default value: -1*
-   **sout-mosaic-bridge-y \** : Y coordinate of the upper left corner in the mosaic if non-negative. *default value: -1*

#### Source code

-   [modules/stream_out/mosaic-bridge.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/stream_out/mosaic-bridge.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### motion control {#modules-motion-control}

Module: motion

**Type**: Control interface

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: motion control interface

**Shortcut(s)**: -

Use this control interface to rotate the video when moving a laptop using HDAPS or AMS sensors.

#### Options

-   **motion-use-rotate** : Use the rotate video filter instead of the transform video ouput filter to rotate the video *default value: enabled*

#### Examples

Using the transform video output filter (possible rotation angles are -90°, 0° and +90°):

    % vlc --control motion somevideo.avi

Using the rotate video filter (possible rotation angles theoretically range from -180° to +180°, depending on your sensor):

    % vlc --control motion --motion-use-rotate --video-filter rotate somevideo.avi

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### motionblur {#modules-motionblur}

Module: motionblur

**Type**: Video filter

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: motion blur video filter

**Shortcut(s)**: -

Use this filter to blur motion in a video based on previous frames.

#### Options

-   **motionblur-factor \** : The bluring factor (1 to 127). Higher values mean more blurring *default value: 80*

#### Examples

    % vlc --video-filter "motionblur{factor=60}" somevideo.avi

**Note:** In versions prior to 0.9.0, motionblur was a video output filter.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### mp4 {#modules-mp4}

Support for fragmented MP4 muxing/demuxing was added in VLC 3.0.0.

The demux module is planned to support [HEIF](http://en.wikipedia.org/wiki/High_Efficiency_Image_File_Format) in future versions (currently in 4.0.0-dev) through a 0 submodule.

#### Demux

Module: mp4

**Type**: Access demux

**First VLC version**: 0.5.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: MP4 stream demuxer

**Shortcut(s)**: (none)

-   **mp4-m4a-audioonly \** : Ignore non audio tracks from iTunes audio files *default value: disabled*

#### Mux

Module: mp4

**Type**: Muxer

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: MP4/MOV muxer

**Shortcut(s)**: 0, 1, 2

-   **sout-mp4-faststart \** : Create "Fast Start" files. "Fast Start" files are optimized for downloads and allow the user to start previewing the file while it is downloading *default value: enabled*

##### mp4frag

Module: mp4frag

**Type**: Muxer

**First VLC version**: 3.0.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Fragmented and streamable MP4 muxer

**Shortcut(s)**: 0, 1

#### Source code

-   [modules/demux/mp4/mp4.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/mp4/mp4.c)
-   [modules/mux/mp4/mp4.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/mux/mp4/mp4.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### mpc {#modules-mpc}

Module: mpc

**Type**: Access demux

**First VLC version**: 0.8.4

**Last VLC version**: -

**Operating system(s)**: all

**Description**: MusePack demuxer

**Shortcut(s)**: 0

The option 0 was removed entirely in [\[7e29d932257d0bf6dca42ffecf9b0dce523ca92e\]](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=7e29d932257d0bf6dca42ffecf9b0dce523ca92e) (0.8.6).

#### Options

None.

#### Source code

-   [modules/demux/mpc.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/mpc.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### mpjpeg {#modules-mpjpeg}

Module: mpjpeg

**Type**: Muxer

**First VLC version**: 0.8.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Multipart JPEG muxer

**Shortcut(s)**: 0

The option 0 was deprecated prior to VLC 0.9.0 with [this commitdiff](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=fc16feb13f52733cfe8c1b1219b519158a4c19c3) to close [Bug #1188](https://trac.videolan.org/vlc/ticket/1188).

#### Options

None.

#### Source code

-   [modules/mux/mpjpeg.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/mux/mpjpeg.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### mqtt {#modules-mqtt}

**NOTE: this module is in active development and has not made it into the main tree yet.**

Module: mqtt

**Type**: Interface

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: Any that support the [mosquitto](http://mosquitto.org/) library

**Description**: control VLC using the MQTT protocol

**Shortcut(s)**: -

This module will let you send control messages to VLC using the [MQTT](http://www.mqtt.org/) protocol.

#### Options

-   **mqtt-host \** : Hostname of MQTT broker to connect to *default value: localhost*
-   **mqtt-port \** : Port number of MQTT broker to connect to *default value: 1883*
-   **mqtt-username \** : The username to connect to the broker with *default value: none*
-   **mqtt-password \** : The password to connect to the broker with *default value: none*
-   **mqtt-prefix \** : The topic name prefix to use *default value: vlc/*
-   **mqtt-clientid \** : The client identifier to connect to the broker as *default value: random*
-   **mqtt-keepalive \** : The keep alive time for the MQTT protocol (in seconds) *default value: 10*
-   **mqtt-qos \** : The QoS level to publish and subscribe using (0, 1 or 2) *default value: 1*

#### Protocol

##### Commands

Sent from client to VLC.

-   **Direction**: Topic Payload Description
-   **→**: vlc/command \ \ Any of the below
-   **→**: vlc/command/add \ Add \ to the playlist
-   **→**: vlc/command/delete \ delete item \ in playlist
-   **→**: vlc/command/clear clear the playlist
-   **→**: vlc/command/play \ Start playing item at
-   **→**: vlc/command/pause Pause Playback
-   **→**: vlc/command/stop Stop Playback
-   **→**: vlc/command/goto \ Goto item at index
-   **→**: vlc/command/next Start playing next item in playlist
-   **→**: vlc/command/prev Start playing prev item in playlist
-   **→**: vlc/command/seek Seek to in the current item (in seconds)
-   **→**: vlc/command/volume \ Set volume to \ (0 to 255)
-   **→**: vlc/command/volup \ Increase volume by
-   **→**: vlc/command/voldown \ Decrease volume by
-   **→**: vlc/command/repeat \ Turn on or off playlist *repeat* mode (0 or 1)
-   **→**: vlc/command/random \ Turn on or off playlist *random* mode (0 or 1)
-   **→**: vlc/command/loop \ Turn on or off playlist *loop* mode (0 or 1)

##### Status

Sent from VLC to client.

-   **Direction**: Topic Payload Description
-   **←**: vlc/status/playlist \ A JSON representation of the playlist is sent whenever the playlist changes.
-   **←**: vlc/status/state \ This retained message is sent by VLC whenever the player changes state:
    -   opening
    -   buffering
    -   playing
    -   paused
    -   stopped
    -   ended
    -   error
    -   notconnected
-   **←**: vlc/status/playing \ Information about the currently playing item as JSON is sent whenever a new item starts playing.
-   **←**: vlc/status/time Progress through the current stream as decimal seconds
-   **←**: vlc/status/length Duration of current stream as decimal seconds
-   **←**: vlc/status/volume \ The current volume between 0 and 255 (inclusive)

### ncurses {#modules-ncurses}

Module: ncurses

**Type**: Control interface

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: ncurses console interface

**Shortcut(s)**: -

#### Introduction

This is one of the three command line interfaces (besides remote control (rc) and telnet). To force vlc into using this interface, do the following:

    vlc -I ncurses

This interface is operated through a set of shortcuts which are listed in the next section.

#### Shortcuts

To get the following list of all available shortcuts in the interface press 'h'. Use the up and down arrow keys to scroll.

    [Display]
    h,H         Show/Hide help box
    i           Show/Hide info box
    L           Show/Hide messages box
    P           Show/Hide playlist box
    B           Show/Hide filebrowser

    [Global]
    q, Q        Quit
    s           Stop
         Pause/Play
    f           Toggle Fullscreen
    n, p        Next/Previous playlist item
    [, ]        Next/Previous title
    <, >        Next/Previous chapter
         Seek +1%
          Seek -1%
    a           Volume Up
    z           Volume Down

    [Playlist]
    r           Random
    l           Loop Playlist
    R           Repeat item
    o           Order Playlist by title
    O           Reverse order Playlist by title
    /           Look for an item
    A           Add an entry
    D,     Delete an entry
     Delete an entry

    [Filebrowser]
         Add the selected file to the playlist
         Add the selected directory to the playlist
    .           Show/Hide hidden files

    [Boxes]
    ,     Navigate through the box line by line
    , Navigate through the box page by page

    [Player]
    ,     Seek +/-5%

    [Miscellaneous]
    Ctrl-l          Refresh the screen

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### netsync {#modules-netsync}

Module: netsync

**Type**: Video output

**First VLC version**: 0.8.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Synchronise remote VLC instances

**Shortcut(s)**: -

#### Introduction

Use this module to keep several clients synchronised on a single VLC stream.

Common uses of this module are:

-   Synchronising lots of loud PC speakers during a party;
-   Synchronising several computers playing parts of a video wall.

#### Options

-   **netsync-master** : Act as master *default value: disabled*
-   **netsync-master-ip \** : Master client ip address *default value: ""*

#### Examples

Here's a small example:

We're going to be listening to a multicast stream.

Run a client as master syncronisation client (master has IP address 192.168.0.1):

    % vlc udp://@239.255.1.1 --control netsync --netsync-master

And on the other clients:

    % vlc udp://@239.255.1.1 --control netsync --netsync-master-ip 192.168.0.1

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### noise {#modules-noise}

Module: noise

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: 1.1.13

**Operating system(s)**: all

**Description**: add random noise to the video

**Shortcut(s)**: -

#### Example

    % vlc --video-filter noise somevideo.avi

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### nsv {#modules-nsv}

Module: nsv

**Type**: Access demux

**First VLC version**: 0.7.1

**Last VLC version**: -

**Operating system(s)**: all

**Description**: NullSoft Video demuxer

**Shortcut(s)**: -

The shortcut for this module is 0.

#### Options

None.

#### Source code

-   [modules/demux/nsv.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/nsv.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### ogg {#modules-ogg}

The ogg demux module refers to Ogg as *OGG*. The Xiph Wiki [clarifies](https://wiki.xiph.org/Ogg#Name) that the name is not an acronym and should be written *Ogg* or *ogg*.

The earliest mention of Ogg muxing support in the changelog was for the macOS port in 0.5.3.

#### Demux

Module: ogg

**Type**: Access demux

**First VLC version**: 0.5.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: OGG demuxer

**Shortcut(s)**: 0

##### Options

None.

#### Mux

Module: ogg

**Type**: Muxer

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Ogg/OGM muxer

**Shortcut(s)**: 0, 1

##### Options

-   **sout-ogg-indexintvl \<integer \[0 .. 0{.variable}\]\>** : Minimal index interval, in [milliseconds](http://en.wiktionary.org/wiki/ms). Set to 0 to disable index creation. *default value: 1000*
-   **sout-ogg-indexratio \** : Set index size ratio. Alters default (60min content) or estimated size. *default value: 1.0*

#### Source code

-   [modules/demux/ogg.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/ogg.c)
-   [modules/demux/ogg_granule.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/ogg_granule.c)
-   [modules/demux/oggseek.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/oggseek.c)
-   [modules/mux/ogg.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/mux/ogg.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### oldmovie {#modules-oldmovie}

Module: oldmovie

**Type**: Video filter

**First VLC version**: 2.2.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Old movie effect video filter

**Shortcut(s)**: -

#### Options

None.

#### Examples

    $ vlc --video-filter "oldmovie" video.ogv

#### Source code

-   [modules/video_filter/oldmovie.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/oldmovie.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### opengl {#modules-opengl}

This module features conditional compilation: [add_module](https://www.videolan.org/developers/vlc/doc/doxygen/html/vlc__plugin_8h.html#a789d7743e2a12bcaef2f0677a81c5c44) monitors 0{.variable} to determine the 1{.sample}.

More simply, 0 is called for desktop computers and 1 is called for embedded devices (e.g. smartphones).

Neither conditional module accepts options—0 is called in [vout_helper.h](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_output/opengl/vout_helper.h) for that purpose.

#### gles2

Module: gles2

**Type**: Video output

**First VLC version**: 2.0.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: OpenGL for Embedded Systems 2 video output

**Shortcut(s)**: 0, 1

#### gl

Module: gl

**Type**: Video output

**First VLC version**: 0.8.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: OpenGL video output

**Shortcut(s)**: 0, 1

OpenGL (as a plugin) was first introduced for the macOS port in VLC 0.7.1, made the default for macOS in VLC 0.7.2, and later added for all platforms in VLC 0.8.0.

#### Source code

-   Git: [modules/video_output/opengl/display.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_output/opengl/display.c) (main file)
-   Git: [modules/video_output/opengl/vout_helper.h](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_output/opengl/vout_helper.h) (OpenGL and OpenGL ES output common code)
-   Doxygen: [include/vlc_opengl.h](https://www.videolan.org/developers/vlc/doc/doxygen/html/vlc__opengl_8h_source.html)
-   Doxygen: [include/vlc_opengl.c](https://www.videolan.org/developers/vlc/doc/doxygen/html/opengl_8c.html)
-   Doxygen: [include/vlc_vout_display.h](https://www.videolan.org/developers/vlc/doc/doxygen/html/vlc__vout__display_8h.html)

#### See also

-   opengl
-   glwin32 - module for Windows 32-bit OpenGL
-   glx - module for Linux X11 OpenGL

#### External links

-   [opengl.org](https://opengl.org/)
-   [www.khronos.org/opengles](https://www.khronos.org/opengles) - developer page for OpenGL ES
-   Wikibook: [OpenGL Programming](http://en.wikibooks.org/wiki/OpenGL_Programming)
    -   chapter [OpenGL ES Overview](http://en.wikibooks.org/wiki/OpenGL_Programming/OpenGL_ES_Overview)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### opensles {#modules-opensles}

Module: opensles

**Type**: Audio output

**First VLC version**: 2.1

**Last VLC version**: -

**Operating system(s)**: Android

**Description**: OpenSLES audio output

**Shortcut(s)**: -

#### Introduction

This is the latest audio output module found on the Android platform. It offers better lipsync than the AudioTrack modules.

Sometimes, however, on recent Android platforms an [Android bug](https://trac.videolan.org/vlc/ticket/9325) might prevent it from functioning well, so in that case trying out the AudioTrack (Java) audio output may be recommended.

#### Options

None.

#### Source code

-   [modules/audio_output/opensles_android.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/audio_output/opensles_android.c)

### osc {#modules-osc}

**NOTE: this module is was never merged into the main VLC codebase and work on it has stopped**

-   Patch against VLC 0.9.9a: 0
-   Patch against VLC 0.9.8a: 0
-   Patch against GIT master: 0

### oss {#modules-oss}

oss and alsa audio capture support were removed from v4l and v4l2 in VLC 1.0.0, but accesses were provided as sub-modules. To emulate old behaviour, use 0 or 1. The access module reads from 2.

#### Options

##### Audio output

Module: oss

**Type**: Audio output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: Open Sound System audio output

**Shortcut(s)**: (none)

-   **oss-audio-device \** : OSS device node path *default value: ""*
-   **oss-spdif \** : S/PDIF can be used by default when your hardware supports it as well as the audio stream being played *default value: disabled*

##### Access

Module: oss

**Type**: Access

**First VLC version**: 1.0.0

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: OSS input

**Shortcut(s)**: 0

-   **oss-stereo \** : Capture the audio stream in stereo *default value: enabled*
-   **oss-samplerate \** : Sample rate of the captured audio stream, in Hz (eg: 11025, 22050, 44100, 48000) *default value: 48000*

#### Source code

-   [modules/audio_output/oss.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/audio_output/oss.c)
-   [modules/access/oss.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/oss.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### panoramix {#modules-panoramix}

Module: Panoramix

**Type**: Video output splitter

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: wall with overlap video filter

**Shortcut(s)**: -

    Panoramix: wall with overlap video filter
    Split the video in multiple windows to display on a wall of screens
         --panoramix-cols= Number of columns
         --panoramix-rows= Number of rows
         --panoramix-bz-length=
                                    length of the overlapping area (in %)
         --panoramix-bz-height=
                                    height of the overlapping area (in %)
         --panoramix-attenuate, --no-panoramix-attenuate
                                    Attenuation (default enabled)
         --panoramix-bz-begin=
                                    Attenuation, begin (in %)
         --panoramix-bz-middle=
                                    Attenuation, middle (in %)
         --panoramix-bz-end=
                                    Attenuation, end (in %)
         --panoramix-bz-middle-pos=
                                    middle position (in %)
         --panoramix-bz-gamma-red=
                                    Gamma (Red) correction
         --panoramix-bz-gamma-green=
                                    Gamma (Green) correction
         --panoramix-bz-gamma-blue=
                                    Gamma (Blue) correction
         --panoramix-bz-blackcrush-red=
                                    Black Crush for Red
         --panoramix-bz-blackcrush-green=
                                    Black Crush for Green
         --panoramix-bz-blackcrush-blue=
                                    Black Crush for Blue
         --panoramix-bz-whitecrush-red=
                                    White Crush for Red
         --panoramix-bz-whitecrush-green=
                                    White Crush for Green
         --panoramix-bz-whitecrush-blue=
                                    White Crush for Blue
         --panoramix-bz-blacklevel-red=
                                    Black Level for Red
         --panoramix-bz-blacklevel-green=
                                    Black Level for Green
         --panoramix-bz-blacklevel-blue=
                                    Black Level for Blue
         --panoramix-bz-whitelevel-red=
                                    White Level for Red
         --panoramix-bz-whitelevel-green=
                                    White Level for Green
         --panoramix-bz-whitelevel-blue=
                                    White Level for Blue
         --panoramix-active=
                                    Active windows

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### playlist {#modules-playlist}

Module: playlist

**Type**: Access demux

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Playlist

**Shortcut(s)**: -

The option 0 has been obsolete since VLC 1.1.0.

#### Options

-   **playlist-skip-ads \** : Use playlist options usually used to prevent ads skipping to detect ads and prevent adding them to the playlist *default value: enabled*

#### Sub-modules

-   **Submodule name**: Description Shortcuts Opens file extensions First version
-   **m3u**: M3U playlist import m3u, m3u8 .m3u, .m3u8, .vlc 0.5.0, 1.0.4
-   **ram**: RAM playlist import N/A .ram, .rm 2.0.2
-   **pls**: PLS playlist import N/A .pls 0.5.2
-   **b4s**: B4S playlist import shout-b4s .b4s 0.6.1
-   **dvb**: DVB playlist import dvb .conf ?
-   **podcast**: Podcast parser podcast N/A 0.8.5
-   **xspf**: XSPF playlist import N/A .xspf 0.8.5
-   **asx**: ASX playlist import N/A .asx, .wax, .wvx 0.5.0
-   **sgimb**: Kasenna MediaBase parser sgimb N/A ?
-   **qtl**: QuickTime Media Link importer qtl .qtl ?
-   **ifo**: Dummy IFO demux N/A .IFO ?
-   **bdmv**: Dummy BDMV demux N/A .BDMV ?
-   **itml**: iTunes Music Library importer itml .xml ?
-   **wpl**: WPL playlist import wpl .wpl, .zpl 1.1.0

#### Source code

-   [modules/demux/playlist/playlist.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/playlist/playlist.c) (file)
-   [modules/demux/playlist](https://git.videolan.org/?p=vlc.git;a=tree;f=modules/demux/playlist;hb=HEAD) (folder)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### podcast {#modules-podcast}

Module: podcast

**Type**: Services discovery

**First VLC version**: 0.8.5

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Podcasts

**Shortcut(s)**: -

#### Options

-   **podcast-urls \** : Enter the list of podcasts to retrieve, separated by '\|' (pipe)

#### Source code

-   [modules/services_discovery/podcast.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/services_discovery/podcast.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### podcast sd {#modules-podcast-sd}

Module: podcast

**Type**: Services discovery

**First VLC version**: 0.8.5

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Podcasts

**Shortcut(s)**: -

#### Options

-   **podcast-urls \** : Enter the list of podcasts to retrieve, separated by '\|' (pipe)

#### Source code

-   [modules/services_discovery/podcast.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/services_discovery/podcast.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### portaudio {#modules-portaudio}

Module: portaudio

**Type**: Audio output

**First VLC version**: 0.8

**Last VLC version**: 2.0

**Operating system(s)**: all

**Description**: Audio output based on the portaudio library (v19)

**Shortcut(s)**: -

**This page is obsolete and kept only for historical interest.** It may document features that are obsolete, superseded, or irrelevant. Do not rely on the information here being up-to-date.

#### Introduction

This was an audio output plugin that used the cross-platform portaudio library to render audio on all platforms.

It was removed in VLC 2.0 Twoflower due to serious problems such as a dependency on the old aout packet API.[\[1\]](https://mailman.videolan.org/pipermail/vlc-devel/2012-January/085344.html) It also had a clock resolution of 1 second, making it impossible for VLC to keep reasonable synchronization with such low precision. Instead of resampling it mostly will discard samples or insert silences.

The Win32 backend of PortAudio was also extremely buggy.

Users should use the waveout plugin for \<= Windows XP, mmdevice (WASAPI) for Windows Vista+, auhal for macOS and pulse (PulseAudio) for Linux.

#### Options

None.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### posterize {#modules-posterize}

Module: posterize

**Type**: Video filter

**First VLC version**: 2.0.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Posterize video by lowering the number of colors

**Shortcut(s)**: -

#### Options

-   **posterize-level \** : Posterize level (number of colors is cube of this value) *default value: 6*

#### Source code

-   [modules/video_filter/posterize.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/posterize.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### projectm {#modules-projectm}

Module: projectm

**Type**: Visualization

**First VLC version**: 1.1.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: libprojectM effect

**Shortcut(s)**: -

Additional module options 0, 1 and 2 [will not be used](http://en.wikipedia.org/wiki/Conditional_compilation) if VLC can find font paths.

#### Options

-   **projectm-width \** : The width of the video window, in pixels *default value: 800*
-   **projectm-height \** : The height of the video window, in pixels *default value: 500*
-   **projectm-meshx \** : The width of the mesh, in pixels *default value: 32*
-   **projectm-meshy \** : The height of the mesh, in pixels *default value: 24*
-   **projectm-texture-size \** : The size of the texture, in pixels *default value: 1024*

#### Source code

-   [modules/visualization/projectm.cpp](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/visualization/projectm.cpp)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### psychedelic {#modules-psychedelic}

Module: psychedelic

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Simulates being high

**Shortcut(s)**: -

#### Options

None.

#### Examples

    % vlc --video-filter "psychedelic" somevideo.avi

**Note:** In versions prior to 0.9.0, this was part of the distort video filter.

#### Source code

-   [modules/video_filter/psychedelic.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/psychedelic.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### pulse {#modules-pulse}

Module: pulse

**Type**: Audio output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: [PulseAudio](http://en.wikipedia.org/wiki/PulseAudio) audio output

**Shortcut(s)**: -

Shortcuts to this module include 0 and 1. Basic PulseAudio *input* support was added in VLC 2.0.0.

#### Introduction

PulseAudio is a sound server associated mainly with GNU/Linux users, but it can also be used on \*BSD and macOS. The pulse and jack modules are two modern options (there might be others) for audio output on Linux. The esd and arts modules were removed prior to the release of VLC 1.0.0.

#### Options

None.

#### Source code

-   [modules/audio_output/pulse.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/audio_output/pulse.c)
-   [modules/audio_output/vlcpulse.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/audio_output/vlcpulse.c) (separate module, support library for libVLC plugins)
-   [modules/access/pulse.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/pulse.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### puzzle {#modules-puzzle}

Module: puzzle

**Type**: Video output filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Turns the video in a jigsaw puzzle game

**Shortcut(s)**: -

#### Options

Note that the puzzle module has been improved in later versions; the option 0 has been **removed** in favour of 1 (use 1 for 2 for the same effect).

-   **puzzle-cols \** : Specifies the number of columns in the puzzle *default value: 4*
-   **puzzle-rows \** : Specifies the number of rows in the puzzle *default value: 4*
-   **puzzle-border \** : Border
-   **puzzle-preview \** : Small preview *default value: disabled*
-   **puzzle-preview-size \** : Small preview size
-   **puzzle-shape-size \** : Piece edge shape size
-   **puzzle-auto-shuffle \** : Puzzle auto shuffle
-   **puzzle-auto-solve \** : Auto solve
-   **puzzle-rotation \** : 0 is (0), 1 is (0/180), 2 is (0/90/180/270), 3 is (0,90,180,270,mirror)
-   **puzzle-mode \** : 0 is jigsaw puzzle, 1 is sliding puzzle, 2 is swap puzzle, 3 is exchange puzzle
-   **puzzle-black-slot** : Change puzzle type to a sliding tile puzzle *default value: disabled*

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### pva {#modules-pva}

Module: pva

**Type**: Access demux

**First VLC version**: 0.7.1

**Last VLC version**: -

**Operating system(s)**: all

**Description**: PVA demuxer

**Shortcut(s)**: 0

#### Options

None.

#### Source code

-   [modules/demux/pva.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/pva.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### pvr {#modules-pvr}

Module: pvr

**Type**: Access

**First VLC version**: -

**Last VLC version**: 2.0.9

**Operating system(s)**: Linux

**Description**: IVTV MPEG Encoding cards input

**Shortcut(s)**: 0

This module was removed with commitdiff [\[fb1dcab36e7c3a49d49562334045e4f87980ac03\]](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=fb1dcab36e7c3a49d49562334045e4f87980ac03).
The changelog for 2.1.0 notes under the section *Removed modules*:

     * PVR: IVTV analog TV encoder - use V4L instead

#### Options

The module did not accept 0 beyond the endpoints given by 1: 23{.variable}4. This was not mentioned in the help text.

The variables in 0 are defined in [modules/access/v4l2/linux/videodev2.h](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/v4l2/linux/videodev2.h).

-   **pvr-device \** : [PVR](http://en.wikipedia.org/wiki/Personal_video_recorder) video device *default value: "/dev/video0"*
-   **pvr-radio-device \** : PVR radio device *default value: "/dev/radio0"*
-   **pvr-norm \ {0{.variable},1{.variable},2{.variable},3{.variable}}** : Norm of the stream (Automatic, SECAM, PAL, or NTSC) *default value: 4{.variable}*
-   **pvr-width \** : Width of the stream to capture (-1 for autodetection) *default value: -1*
-   **pvr-height \** : Height of the stream to capture (-1 for autodetection) *default value: -1*
-   **pvr-frequency \** : Frequency to capture (in kHz), if applicable *default value: -1*
-   **pvr-framerate \** : Framerate to capture, if applicable (-1 for autodetect) *default value: -1*
-   **pvr-keyint \** : Interval between keyframes (-1 for autodetect) *default value: -1*
-   **pvr-bframes \** : If this option is set, B-Frames will be used. Use this option to set the number of B-Frames *default value: -1*
-   **pvr-bitrate \** : Bitrate to use (-1 for default) *default value: -1*
-   **pvr-bitrate-peak \** : Peak bitrate in VBR mode *default value: -1*
-   **pvr-bitrate-mode \ {0,1}** : Bitrate mode to use (VBR or CBR) *default value: -1*
-   **pvr-audio-bitmask \** : [Bitmask](http://en.wiktionary.org/wiki/bitmask) that will get used by the audio part of the card *default value: -1*
-   **pvr-audio-volume \** : Audio volume (0-65535) *default value: -1*
-   **pvr-channel \** : Channel of the card to use (Usually: 0 - tuner, 1 - composite, 2 - svideo) *default value: -1*

#### Source code

-   [modules/access/pvr.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/pvr.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Qt4 {#modules-qt4}

This page is outdated and information might be incorrect.

#### Introduction

Qt / Qt4 is the default, plain, graphical, interface to VLC, made using the [Qt](https://www.qt.io) library (Linux users may need to have this installed). It is used as the default interface on the Windows and Linux versions of VLC media player from version 0.9.0 and above.

Unless you change the preferences, VLC will start up in the Qt4 interface, but you can force this by running

    % vlc -I qt

or

    % qvlc

If Qt4 is not avaliable, it will probably revert to using the rc (console) interface, even if you force it. The most likely reason for this is if Qt4 hasn't been installed, or if it wasn't linked in (using the ./configure). See compiling VLC for information on compiling.

Please note that the pre-0.9.0 wxWidgets interface is replaced by the Qt interface and will thus not be further developed.

#### Using

##### Launching Modes

##### Video Modes

##### Other options

#### Main interface Description

##### Menu Bar

##### Status Bar

The Status Bar has three distinctive portions displaying information.

-   Name label
-   Speed label
-   Time label

Here are the actions you can do on the status bar:

-   **Select**: Left Click Right Click Middle Click Double Click
-   **Name label**: Prepare for copy Give focus for arrows selection Copy menu - Word selection
-   **Speed label**: - - Show fine rate speed adjusting - reset to normal 1.00x speed
-   **Time label**: - Switch remaining/elapsed time Switch remaining/elapsed time Switch remaining/elapsed time Open the "Go To Time" dialog

###### Time label

The timeLabel shows mm:ss/nn:tt in the statusBar.

If the time is longer that one hour, it shows hh:mm:ss/ii:nn:tt.

If you click on it, it shows remaining time instead of elapsed time. If you double click on it, it opens the open the GOTOTime dialog in order to skip easily.

If VLC doesn't know the total time, like in AVI/HTTP, it shows the elapsed time only, and --:-- instead of total time.

##### System Tray Icon

##### Controller

###### Main Slider

The main slider does control the timeline.

When you click somewhere on the timeline, it skips the movie to that place. When you drag the slider of the timeline, it follows the position on the movie.

When you hover it with your mouse, it shows you where it would go if you click.

###### Sound Slider

The sound slider does control the volume.

When you click somewhere on the soundSlider, it changes the volume. When you click and drag on the soundSlider, it changes the volume too. If you release the click really outside of the volume slider, it will reset to your old value.

When you hover it with your mouse, it shows you the volume it would be if you click it.

###### Sound range

The sound goes from 0% to 125% (previously the sound went from 0% to 200% and could go up to 400%).

100% means normal output of the file without amplification. Above 100% means that it may use software amplification, and it could distort sound (it usually doesn't, but it could).

##### Video Widget

##### Background Widget

#### Playlist

The playlist display has three views, the Icon view, the Detailed view, and the List view. The view can be changed by clicking on the icon above the playlist window.

The playlist can be cleared or re-sorted by right-clicking on the background of the display pane. It's also possible to change the size of the display from this menu.

#### Dialogs description

##### Open Dialog

##### Sout Dialog

##### Extended Dialog

The Extended settings dialog box can be brought up by selected the *Ex* button on the main window. Alternatively it can be found in *Tools* menu under *Extended settings...* (Ctrl+E). The extended settings allow effects to be enabled and adjusted in realtime whilst media is playing.

###### Audio Effects

###### Graphic Equalizer

The Graphic equalizer enables the equalization of the sound output to be enabled and adjusted. Tick the *Enable* tickbox to enable it. Ticking the *2 pass* will mean the sound is reprocessed through the equalizer again, thus amplifying the changes. Slide the sliders to adjust the output levels for the preamp and each frequency range. There are a number of presets which can be selected from the drop-down box.

###### Spatializer

The Spatializer contains several adjustable audio filters to change the perceived environment such as room size and humidity. Tick the *Enable spatializer* tickbox to enable and adjust the sliders.

###### Video Effects

###### Basic

###### Color fun

###### Some random name

###### Image modification

###### Find a name

###### Overlay

###### Advanced video filter controls

A colon separated list will be generated when filters and effects in the other video tabs are enabled. The list of effect can be modified. Press the *Update* button to apply the changes.

##### VLM Dialog

#### Interface Hotkeys

Qt Interface Hotkeys

#### Work

QtIntfTODO

### qtcapture {#modules-qtcapture}

Module: qtcapture

**Type**: Access

**First VLC version**: 2.0.0

**Last VLC version**: 2.2.8

**Operating system(s)**: macOS

**Description**: Quicktime Capture

**Shortcut(s)**: -

The qtcapture module was removed prior to 3.0.0, and users were directed to avcapture.

#### Options

-   **qtcapture-width \** : Video Capture width in pixel *default value: 640*
-   **qtcapture-height \** : Video Capture height in pixel *default value: 480*

#### Source code

-   [modules/access/qtcapture.m](https://git.videolan.org/?p=vlc/vlc-2.2.git;a=blob;f=modules/access/qtcapture.m) (vlc/vlc-2.2.git)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### qtsound {#modules-qtsound}

Module: qtsound

**Type**: Access

**First VLC version**: 2.0.0

**Last VLC version**: -

**Operating system(s)**: macOS

**Description**: QuickTime Sound Capture

**Shortcut(s)**: -

The qtsound module is planned to be removed in 4.0.0 and replaced with avaudiocapture.

#### Options

None.

#### Source code

-   [modules/access/qtsound.m](https://git.videolan.org/?p=vlc/vlc-3.0.git;a=blob;f=modules/access/qtsound.m) (vlc/vlc-3.0.git)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### rawdv {#modules-rawdv}

Module: rawdv

**Type**: Access

**First VLC version**: 0.5.0

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: DV (Digital Video) demuxer

**Shortcut(s)**: 0

#### Options

-   **rawdv-hurry-up \** : The demuxer will advance timestamps if the input can't keep up with the rate *default value: disabled*

#### Source code

-   [modules/demux/rawdv.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/rawdv.c)
-   [modules/demux/rawdv.h](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/rawdv.h) (helper)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### rawvid {#modules-rawvid}

Module: rawvid

**Type**: Access demux

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Raw video demuxer

**Shortcut(s)**: 0

#### Options

-   **rawvid-fps \** : This is the desired frame rate when playing raw video streams. In the form 30000/1001 or 29.97
-   **rawvid-width \** : This specifies the width in pixels of the raw video stream
-   **rawvid-height \** : This specifies the height in pixels of the raw video stream
-   **rawvid-chroma \** : Force chroma. This is a four character string
-   **rawvid-aspect-ratio \** : Aspect ratio (4:3, 16:9). Default assumes square pixels

#### Source code

-   [modules/demux/rawvid.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/rawvid.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### real {#modules-real}

Module: real

**Type**: Access demux

**First VLC version**: 0.7.1

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Real plugin

**Shortcut(s)**: -

Shortcuts to this module are 0 and 1 (for RealMedia).

The real module is planned to be removed (currently written as 4.0.0-dev).

#### Options

None.

#### Source code

-   [modules/demux/real.c](https://git.videolan.org/?p=vlc/vlc-3.0.git;a=blob;f=modules/demux/real.c) (vlc/vlc-3.0.git)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### record {#modules-record}

Module: record

**Type**: Access filter

**First VLC version**: 0.8.2

**Last VLC version**: -

**Operating system(s)**: all

**Description**: toggle recording incoming data to disk

**Shortcut(s)**: -

This access filter will enable recording incoming data to disk when the user presses the r key. Note that this is very unlikely to work for sources using an encapsulation method other than ts.

#### Options

-   **record-path \** : Directory where recorded that will be stored. *default value: ""*

#### Example

    % vlc --access-filter record

VLC will toggle recording when you press the r hotkey.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### ripple {#modules-ripple}

Module: ripple

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: ripple video filter

**Shortcut(s)**: -

#### Example

**VLC 0.9.0 and above**:

    % vlc --video-filter ripple somevideo.avi

**Note:** In versions prior to 0.9.0, ripple was part of the distort video output filter.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### rotate {#modules-rotate}

Module: rotate

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: rotate video filter

**Shortcut(s)**: -

Use this filter to rotate the video using any angle you want.

#### Options

-   **rotate-angle \** : Rotation angle in degrees (0 to 359) *default value: 0*
-   **rotate-use-motion \** : Use HDAPS, AMS, APPLESMC or UNIMOTION motion sensors to rotate the video *default value: disabled*

#### Example

    % vlc --video-filter "rotate{angle=123}" somevideo.avi

#### See also

-   Documentation:Modules/motion_control

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### rss {#modules-rss}

Module: rss

**Type**: Video sub-filter

**First VLC version**: 0.8.4

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Overlays [RSS](http://en.wikipedia.org/wiki/RSS) and [Atom](http://en.wikipedia.org/wiki/Atom_(Web_standard) "wikipedia:Atom (Web standard)") feeds on the video

**Shortcut(s)**: 0, 1

#### Options

-   **rss-urls \** : Pipe **0** separated list of [RSS](http://en.wikipedia.org/wiki/RSS) and/or [Atom](http://en.wikipedia.org/wiki/Atom_(Web_standard) "wikipedia:Atom (Web standard)") feed URLs *default value: NULL*

##### Position

-   **rss-x \** : X offset, from the left screen edge *default value: 0*
-   **rss-y \** : Y offset, down from the top *default value: 0*
-   **rss-position \ {0, 1, 2, 4, 8, 5, 6, 9, 10}** : You can enforce the text position on the video (0=center, 1=left, 2=right, 4=top, 8=bottom; you can also use combinations of these values, eg 6 = top-right) *default value: -1*

##### Font

-   **rss-opacity \** : Opacity (inverse of transparency) of overlay text. 0 = transparent, 255 = totally opaque *default value: 255*
-   **rss-color \ { 0x000000, 0x808080, 0xC0C0C0, 0xFFFFFF, 0x800000, 0xFF0000, 0xFF00FF, 0xFFFF00, 0x808000, 0x008000, 0x008080, 0x00FF00, 0x800080, 0x000080, 0x0000FF, 0x00FFFF }** : Color^(**key**)^ of the text that will be rendered on the video. This must be an hexadecimal (like HTML colors). The first two chars are for red, then green, then blue *default value: 0xFFFFFF*
-   **rss-size \** : Font size, in pixels. Default is 0 (use default font size) *default value: 0*

##### Misc

-   **rss-speed \** : Speed of the RSS/Atom feeds in [µs](http://en.wiktionary.org/wiki/%C2%B5s) (bigger is slower) *default value: 100000*
-   **rss-length \** : Maximum number of characters displayed on the screen *default value: 60*
-   **rss-ttl \** : Time in seconds between each feed refresh of the feeds. 0 means that the feeds are never updated. 1800 seconds is 30 minutes *default value: 1800*
-   **rss-images \** : Display feed images if available *default value: enabled*
-   **rss-title \ {0{.variable}, 1{.variable}, 2{.variable}, 3{.variable}}** : Title display mode. 0 is hidden if the feed has an image and feed images are enabled, 1 is always visible, 2 is scroll with feed *default value: 4{.variable}*

#### Examples

Example command line use **(VLC 0.9.0 and above)**: (untested with 3.x.x)

    % vlc somevideo.avi --sub-filter=rss --rss-urls="0

#### Source code

-   [modules/spu/rss.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/spu/rss.c)

#### See also

-   Documentation:Modules/marq

#### Appendix

-   \^ --rss-color
-   \^ --rss-position

&nbsp;

-   **Sample**: Colour Hex code
-   **Black**: 0
-   **Gray**: 0
-   **Silver**: 0
-   **White**: 0
-   **Maroon**: 0
-   **Red**: 0
-   **Fuchsia**: 0
-   **Yellow**: 0
-   **Olive**: 0
-   **Green**: 0
-   **Teal**: 0
-   **Lime**: 0
-   **Purple**: 0
-   **Navy**: 0
-   **Blue**: 0
-   **Aqua**: 0

&nbsp;

-   **Integer**: Alignment Comment
-   **0**: Center
-   **1**: Left
-   **2**: Right
-   **4**: Top
-   **8**: Bottom
-   **5**: Top-Left 4 + 1
-   **6**: Top-Right 4 + 2
-   **9**: Bottom-Left 8 + 1
-   **10**: Bottom-Right 8 + 2
-   **3**: n/a contradictory
-   **7**: n/a contradictory

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### rtp {#modules-rtp}

See also: Documentation:Modules/live555

Module: rtp

**Type**: Access

**First VLC version**: 0.7.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Real-Time Protocol (RTP) input

**Shortcut(s)**: 0, 1, 2

The only supported format for 0 is 1.

#### SRTP

The module supports RTP with encryption (SRTP) through [srtp.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/rtp/srtp.c) using [libgcrypt](https://directory.fsf.org/wiki/Libgcrypt) ([gcrypt manual](https://www.gnupg.org/documentation/manuals/gcrypt/)). There are no sub-modules or other shortcuts (in particular, srtp will not work).

Hexadecimal strings are base-16 numbers. Each character is one of 0123456789abcdef (case-insensitive).

##### Crypto

Functions of interest (defined in [srtp.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/rtp/srtp.c) and [srtp.h](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/rtp/srtp.h)) lie in [rtp.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/rtp/rtp.c) between:

    #ifdef HAVE_SRTP

and

    #endif

In summary:

-   SRTP sessions are one-way and re-keyed periodically
-   To set or reset the master key and master salt for an SRTP session 0 is called
-   The 0 values are currently hard-coded as [AES](http://en.wikipedia.org/wiki/Advanced_Encryption_Standard) in [counter mode](http://en.wikipedia.org/wiki/Block_cipher_mode_of_operation#CTR) authenticated with [HMAC](http://en.wikipedia.org/wiki/HMAC)-[SHA1](http://en.wikipedia.org/wiki/SHA1); the salt with [PRF](http://en.wikipedia.org/wiki/Pseudorandom_function_family)-AES-CM. There are code comments suggesting this be improved
    -   [SHA1 is deprecated](https://shattered.io/) but using it here should be passably secure for now
-   There are explanations (for hackers) in the form of code comments in the files

#### Options

-   **rtcp-port \** : RTCP packets will be received on this transport protocol port. If zero, multiplexed RTP/RTCP is used *default value: 0*
-   **srtp-key \** : RTP packets will be authenticated and deciphered with this Secure RTP master shared secret key. This must be a 32-character-long hexadecimal string
-   **srtp-salt \** : Secure RTP requires a (non-secret) master [salt](http://en.wikipedia.org/wiki/salt_(cryptography) "wikipedia:salt (cryptography)") value. This must be a 28-character-long hexadecimal string
-   **rtp-max-src \** : How many distinct active RTP sources are allowed at a time *default value: 1*
-   **rtp-timeout \** : How long to wait (in seconds) for any packet before a source is expired *default value: 5*
-   **rtp-max-dropout \** : RTP packets will be discarded if they are too much ahead (i.e. in the future) by this many packets from the last received packet *default value: 3000*
-   **rtp-max-misorder \** : RTP packets will be discarded if they are too far behind (i.e. in the past) by this many packets from the last received packet *default value: 100*
-   **rtp-dynamic-pt \** : This payload format will be assumed for dynamic payload types (between 96 and 127) if it can't be determined otherwise with out-of-band mappings (SDP) *default value: NULL*

#### Source code

-   [modules/access/rtp/rtp.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/rtp/rtp.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### rtsp {#modules-rtsp}

Module: live555

**Type**: Access demux

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: RTP/RTSP/SDP demuxer (using Live555)

**Shortcut(s)**: 0, 1

The 0 option was removed prior to VLC 2.0.0 with this commitdiff: [Unify (ACCESS\|DEMUX)\_GET_PTS_DELAY](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=31ac20b22fc37bcf78991159bf8a0f138db05b44)

#### Options

None.

#### Submodule

Module: live555

**Type**: Access

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: RTSP/RTP access and demux

**Shortcut(s)**: 0, 1, 2, 3

##### Options

-   **rtsp-tcp \** : Use RTP over RTSP (TCP) *default value: disabled*
-   **rtp-client-port \** : Port to use for the RTP source of the session *default value: -1*
-   **rtsp-mcast \** : Force multicast RTP via RTSP *default value: disabled*
-   **rtsp-http \** : Tunnel RTSP and RTP over HTTP *default value: disabled*
-   **rtsp-http-port \** : Port to use for tunneling the RTSP/RTP over HTTP *default value: 80*
-   **rtsp-kasenna \** : Kasenna servers use an old and nonstandard dialect of RTSP. With this parameter VLC will try this dialect, but then it cannot connect to normal RTSP servers *default value: disabled*
-   **rtsp-wmserver \** : WMServer uses a nonstandard dialect of RTSP. Selecting this parameter will tell VLC to assume some options contrary to [RFC 2326](https://tools.ietf.org/html/rfc2326) guidelines *default value: disabled*
-   **rtsp-user \** : Sets the username for the connection, if no username or password are set in the url *default value: NULL*
-   **rtsp-pwd \** : Sets the password for the connection, if no username or password are set in the url *default value: NULL*
-   **rtsp-frame-buffer-size \** : RTSP start frame buffer size of the video track, can be increased in case of broken pictures due to too small buffer *default value: 250000*

#### Source code

-   [modules/access/live555.cpp](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/live555.cpp)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### sap {#modules-sap}

Module: sap

**Type**: Services discovery

**First VLC version**: 0.8.2

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Network streams (SAP)

**Shortcut(s)**: -

This module will listen on port 9875 for SAP announcements. The option 0 has been deprecated since 1.0.0 (redundant). Options 1 and 2 have been deprecated since 2.0.0.

#### Options

-   **sap-addr \** : The SAP module normally chooses itself the right addresses to listen to. However, you can specify a specific address *default value: NULL*
-   **sap-timeout \** : Delay after which SAP items get deleted if no new announcement is received *default value: 1800*
-   **sap-parse \** : This enables actual parsing of the announces by the SAP module. Otherwise, all announcements are parsed by the "live555" (RTP/RTSP) module *default value: enabled*
-   **sap-strict \** : When this is set, the SAP parser will discard some non-compliant announcements *default value: disabled*

#### Sub-module

Module: sap

**Type**: Services discovery

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: SDP Descriptions parser

**Shortcut(s)**: 0

##### Options

None.

#### Source code

-   [modules/services_discovery/sap.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/services_discovery/sap.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### scale {#modules-scale}

Module: scale

**Type**: Video filter

**First VLC version**: 0.8.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Video scaling filter

**Shortcut(s)**: (none)

This module uses the low quality "nearest neighbour" algorithm.
[ARGB](http://en.wikipedia.org/wiki/ARGB) support [was introduced](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=91106e6a04862979b498f3cc23d14eb2057fbd5d) in VLC 3.0.0 (not mentioned in the source code header).

Supported formats for RGBA colour space:

-   RGBA
-   RGB32
-   ARGB

Supported formats for YUV colour space:

-   I420
-   YV12
-   YUVP
-   YUVA

#### Options

None.

#### Source code

-   [modules/video_filter/scale.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/scale.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### scene {#modules-scene}

Module: scene

**Type**: Video filter

**First VLC version**: 1.0.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Send your video to picture files

**Shortcut(s)**: -

**Note:** Before version 1.0.0, this used to be image.

#### Options

-   **scene-format \** : Image format. Format of the output images (png, jpeg, ...) *default value: png*
-   **scene-width \** : Image width. You can enforce the image width. By default (-1) VLC will adapt to the video characteristics *default value: -1*
-   **scene-height \** : Image height. You can enforce the image height. By default (-1) VLC will adapt to the video characteristics *default value: -1*
-   **scene-prefix \** : Filename prefix. Prefix of the output images filenames. Output filenames will have the "prefixNUMBER.format" form if 0{.variable} is not true *default value: scene*
-   **scene-path \** : Directory path prefix. Directory path where images files should be saved. If not set, then images will be automatically saved in users homedir *default value: NULL*
-   **scene-replace \** : Always write to the same file. Always write to the same file instead of creating one file per image. In this case, the number is not appended to the filename *default value: disabled*
-   **scene-ratio \<integer \[1 .. 0{.variable}\]\>** : Recording ratio. Ratio of images to record. 3 means that one image out of three is recorded *default value: 50*

#### Source code

-   [modules/video_filter/scene.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/scene.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### schroedinger {#modules-schroedinger}

The schroedinger module replaces the earlier dirac module for decoding and encoding Dirac, a video codec.

#### Demux

Module: schroedinger

**Type**: Access demux

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Dirac video decoder using libschroedinger

**Shortcut(s)**: 0

##### Options

None.

#### Mux

Module: schroedinger

**Type**: Muxer

**First VLC version**: 1.1.8

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Dirac video encoder using libschroedinger

**Shortcut(s)**: 0, 1

##### Options

-   **sout-schro-rate-control \ {constant_noise_threshold,constant_bitrate,low_delay,lossless,constant_lambda,constant_error,constant_quality}** : Method used to encode the video sequence *default value: NULL*
-   **sout-schro-quality \** : Quality factor to use in constant quality mode *default value: -1.*
-   **sout-schro-noise-threshold \** : Noise threshold to use in constant noise threshold mode *default value: -1.*
-   **sout-schro-bitrate \<integer \[-1 .. 0{.variable}\]\>** : Target bitrate in kbps when encoding in constant bitrate mode *default value: -1*
-   **sout-schro-max-bitrate \<integer \[-1 .. 0{.variable}\]\>** : Maximum bitrate in kbps when encoding in constant bitrate mode *default value: -1*
-   **sout-schro-min-bitrate \<integer \[-1 .. 0{.variable}\]\>** : Minimum bitrate in kbps when encoding in constant bitrate mode *default value: -1*
-   **sout-schro-gop-structure \ {adaptive,intra_only,backref,chained_backref,biref,chained_biref}** : GOP structure used to encode the video sequence *default value: NULL*
-   **sout-schro-gop-length \<integer \[-1 .. 0{.variable}\]\>** : Number of pictures between successive sequence headers i.e. length of the group of pictures *default value: -1*
-   **sout-schro-chroma-fmt \ {420,422,444}** : Picking chroma format will force a conversion of the video into that format *default value: 420*
-   **sout-schro-coding-mode \ {auto,progressive,field}** : Field coding is where interlaced fields are coded separately as opposed to a pseudo-progressive frame *default value: auto*
-   **sout-schro-mv-precision \ {1,1/2,1/4,1/8}** : Motion Vector precision in pels *default value: NULL*
-   **sout-schro-intra-wavelet \ {desl_dubuc_9_7,le_gall_5_3,desl_dubuc_13_7,haar_0,haar_1,fidelity,daub_9_7}** : Intra picture DWT filter *default value: NULL*
-   **sout-schro-inter-wavelet \ {desl_dubuc_9_7,le_gall_5_3,desl_dubuc_13_7,haar_0,haar_1,fidelity,daub_9_7}** : Inter picture DWT filter *default value: NULL*
-   **sout-schro-transform-depth \<integer \[-1 .. 0{.variable}\]\>** : Also known as DWT levels *default value: -1*
-   **sout-schro-filtering \ {none,center_weighted_median,gaussian,add_noise,adaptive_gaussian,lowpass}** : Enable adaptive prefiltering *default value: NULL*
-   **sout-schro-filter-value \** : Higher value implies more prefiltering *default value: -1.*

##### Advanced options

-   **sout-schro-motion-block-size \ {auto,small,medium,large}** : Size of motion compensation blocks *default value: NULL*
-   **sout-schro-motion-block-overlap \ {automatic,none,partial,full}** : Overlap of motion compensation blocks *default value: NULL*
-   **sout-schro-me-combined \** : Use chroma as part of the motion estimation process *default value: -1*
-   **sout-schro-enable-hierarchical-me \** : Enable hierarchical Motion Estimation *default value: -1*
-   **sout-schro-downsample-levels \** : Number of levels of downsampling in hierarchical motion estimation mode *default value: -1*
-   **sout-schro-enable-global-me \** : Enable Global Motion Estimation *default value: -1*
-   **sout-schro-enable-phasecorr-me \** : Enable Phase Correlation Estimation *default value: -1*
-   **sout-schro-enable-multiquant \** : Enable multiple quantizers per subband (one per codeblock) *default value: -1*
-   **sout-schro-codeblock-size \ {automatic,small,medium,large,full}** : Size of code blocks in each subband *default value: NULL*
-   **sout-schro-enable-scd \** : Enable Scene Change Detection *default value: -1*
-   **sout-schro-perceptual-weighting \ {none,ccir959,moo,manos_sakrison}** : perceptual weighting method *default value: NULL*
-   **sout-schro-perceptual-distance \** : perceptual distance to calculate perceptual weight *default value: -1*
-   **sout-schro-enable-noarith \** : Use variable length codes instead, useful for very high bitrates *default value: -1*
-   **sout-schro-horiz-slices \<integer \[-1 .. 0{.variable}\]\>** : Number of horizontal slices per frame in low delay mode *default value: -1*
-   **sout-schro-vert-slices \<integer \[-1 .. 0{.variable}\]\>** : Number of vertical slices per frame in low delay mode *default value: -1*
-   **sout-schro-force-profile \ {auto,vc2_low_delay,vc2_simple,vc2_main,main}** : Force Profile *default value: NULL*

##### Appendix

**For the option 0:**

adaptive
:   No fixed GOP structure. A picture can be intra or inter and refer to previous or future pictures.

intra_only
:   I-frame only sequence

backref
:   Inter pictures refere to previous pictures only

chained_backref
:   Inter pictures refere to previous pictures only

biref
:   Inter pictures can refer to previous or future pictures

chained_biref
:   Inter pictures can refer to previous or future pictures

------------------------------------------------------------------------

**For the option 0:**

420
:   4:2:0

422
:   4:2:2

444
:   4:4:4

------------------------------------------------------------------------

**For the option 0:**

auto
:   auto - let encoder decide based upon input (Best)

progressive
:   force coding frame as single picture

field
:   force coding frame as separate interlaced fields

------------------------------------------------------------------------

**For the option 0:**

automatic
:   automatic - let encoder decide based upon input (Best)

small
:   small - use small motion compensation blocks

medium
:   medium - use medium motion compensation blocks

large
:   large - use large motion compensation blocks

------------------------------------------------------------------------

**For the option 0:**

automatic
:   automatic - let encoder decide based upon input (Best)

none
:   none - Motion compensation blocks do not overlap

partial
:   partial - Motion compensation blocks only partially overlap

full
:   full - Motion compensation blocks fully overlap

------------------------------------------------------------------------

**For the options 0 and 1:**

desl_dubuc_9_7
:   Deslauriers-Dubuc (9,7)

le_gall_5_3
:   LeGall (5,3)

desl_dubuc_13_7
:   Deslauriers-Dubuc (13,7)

haar_0
:   Haar with no shift

haar_1
:   Haar with single shift per level

fidelity
:   Fidelity filter

daub_9_7
:   Daubechies (9,7) integer approximation

------------------------------------------------------------------------

**For the option 0:**

automatic
:   automatic - let encoder decide based upon input (Best)

small
:   small - use small code blocks

medium
:   medium - use medium sized code blocks

large
:   large - use large code blocks

full
:   full - One code block per subband

------------------------------------------------------------------------

**For the option 0:**

none
:   No pre-filtering

center_weighted_median
:   Centre Weighted Median

gaussian
:   Gaussian Low Pass Filter

add_noise
:   Add Noise

adaptive_gaussian
:   Gaussian Adaptive Low Pass Filter

lowpass
:   Low Pass Filter

------------------------------------------------------------------------

**For the option 0:**

auto
:   automatic - let encoder decide based upon input (Best)

vc2_low_delay
:   VC2 Low Delay Profile

vc2_simple
:   VC2 Simple Profile

vc2_main
:   VC2 Main Profile

main
:   Main Profile

#### Source code

-   [modules/codec/schroedinger.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/schroedinger.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### screen {#modules-screen}

Module: screen

**Type**: Access

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: screen capture

**Shortcut(s)**: -

Stream or save a video of your computer screen.

#### Options

-   **screen-caching \** : Time in milliseconds

&nbsp;

-   **screen-fps \** : Capture frames per second *default value: 0*

&nbsp;

-   **screen-fragment-size \** : (Windows only) Optimize the capture by fragmenting the screen in chunks of predefined height (16 might be a good value, and 0 means disabled) *default value: 0*

&nbsp;

-   **screen-top \** : The top edge coordinate of the subscreen. (New in VLC 0.9.0 on x11, New in VLC 1.0.0 on Windows) *default value: 0*

&nbsp;

-   **screen-left \** : The left edge coordinate of the subscreen. (New in VLC 0.9.0 on x11, New in VLC 1.0.0 on Windows) *default value: 0*

&nbsp;

-   **screen-width \** : The width of the subscreen. (New in VLC 0.9.0 on x11, New in VLC 1.0.0 on Windows) *default value: \*

&nbsp;

-   **screen-height \** : The height of the subscreen. (New in VLC 0.9.0 on x11, New in VLC 1.0.0 on Windows) *default value: \*

&nbsp;

-   **screen-follow-mouse, no-screen-follow-mouse** : Follow the mouse when capturing a subscreen. (New in VLC 0.9.0 on x11, New in VLC 1.0.0 on Windows) *default value: no-screen-follow-mouse*

&nbsp;

-   **screen-mouse-image \** : (Windows only) Mouse pointer image to use. If specified, the pointer will be overlayed on the captured video. (New in VLC 1.0.0) *default value: ""*

**screen-mouse-image notes:** - The registration point is (at least by defualt) at the top left corner of image. - File location is relative to your VLC installation folder

Run...

    % vlc -H

...for the definitive options for your version.

#### Example

Capture a screen:

    % vlc screen:// --screen-fps=1 --screen-width=100 --screen-height=100

The screen thus captured is 100x100 pixels in from the top left corner of the active screen.

Show mouse on screen:

    % vlc screen:// --screen-fps=30 :screen-mouse-image=file:///c:/cursorimage.png

##### Questions

How to save? Where is the file saved?

Example of a file save (:file{dst=D:\\\\savedir.mp4}):

    % vlc screen:// :sout=#transcode{vcodec=h264,vb=0,scale=0,acodec=mpga,ab=128,channels=2,samplerate=44100}:file{dst=D:\\savedir.mp4} :sout-keep :screen-mouse-image=file:///c:/cursorimage.png

How to get audio to work?

On a dual head monitor, how to make sure the recording is happening on target monitor?

##### Commands I've used

vlc screen:// --dshow-fps=29.950001 --nooverlay --sout #transcode{vcodec=h264,vb=800, scale=0.5,acodec=mp3,ab=128,channels=2} :duplicate{dst=std{access=file, mux=mp4,dst=/home/user/Desktop/test.flv}}

Produced a black screen... on my Fedora 12 machine.

#### Screenshot

The following screenshot is unrelated to the previous demonstration.

One user created a Droste effect with a screen feed of the screen feed. Click the image to view full-size.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### sdl aout {#modules-sdl-aout}

Module: SDL

**Type**: Audio output

**First VLC version**: -

**Last VLC version**: 1.1.13

**Operating system(s)**: all

**Description**: [Simple DirectMedia Layer](http://en.wikipedia.org/wiki/Simple_DirectMedia_Layer) audio output

**Shortcut(s)**: -

**This page is obsolete and kept only for historical interest.** It may document features that are obsolete, superseded, or irrelevant. Do not rely on the information here being up-to-date.

This module had a single shortcut: 0.

#### Options

None

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### sdl vout {#modules-sdl-vout}

Module: sdl

**Type**: Video output

**First VLC version**: -

**Last VLC version**: 2.2.8

**Operating system(s)**: all

**Description**: [SDL](http://en.wikipedia.org/wiki/Simple_DirectMedia_Layer) video output

**Shortcut(s)**: -

**This page is obsolete and kept only for historical interest.** It may document features that are obsolete, superseded, or irrelevant. Do not rely on the information here being up-to-date.

This module had a single shortcut: 0. Information on this page was [adapted](https://git.videolan.org/?p=vlc/vlc-1.1.git;a=blob;f=modules/video_output/sdl.c;h=beb01eff60081b4b1e8f6872a132fa30ee21359b;hb=HEAD) from the 1.1 branch of vlc.git. 1 was marked as deprecated since 1.1.0.

#### Options

-   **sdl-chroma \** : Force the SDL renderer to use a specific chroma format instead of trying to improve performances by using the most efficient one
-   **sdl-video-driver \** : Force a specific SDL video output driver

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### sdp {#modules-sdp}

Module: sdp

**Type**: Access

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Session Description Protocol: Fake input for 0 scheme

**Shortcut(s)**: 0

#### Options

None.

#### Source code

-   [modules/access/sdp.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/sdp.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### sepia {#modules-sepia}

Module: sepia

**Type**: Video filter

**First VLC version**: 2.0.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Gives video a warmer tone by applying sepia effect

**Shortcut(s)**: -

#### Options

-   **sepia-intensity \** : Intensity of sepia effect *default value: 120*

#### Examples

    % --video-filter "sepia" video.ogv
    % --video-filter "sepia{intensity=100}" video.ogv

#### Source code

-   [modules/video_filter/sepia.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/sepia.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### sharpen {#modules-sharpen}

Module: sharpen

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: sharpening video filter

**Shortcut(s)**: -

Use this filter to sharpen the video.

#### Options

-   **sharpen-sigma \** : Sharpen strength (0. to 2.) *default value: 0.05*

#### Example

    % vlc --video-filter "sharpen{sigma=0.12}" somevideo.avi

#### See also

-   Documentation:Modules/gaussianblur

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### shout {#modules-shout}

The services discovery module was removed. **The Access output module is current.**

#### Access output

Module: shout

**Type**: Access output

**First VLC version**: 0.8.4

**Last VLC version**: -

**Operating system(s)**: all

**Description**: This module forwards vorbis streams to an icecast server

**Shortcut(s)**: 0

Documentation is present directly in the source code (4.0.0-dev) as multiple C comment blocks, relevant comments reproduced here (copyright © 2005 VLC authors and VideoLAN, Authors: Daniel Fischer and Derk-Jan Hartman, LGPL 2.1 or later):

    /*
     * Some Comments:
     *
     * - this only works for ogg and/or mp3, and we don't check this yet.
     * - MP3 metadata is not passed along, since metadata is only available after
     *   this module is opened.
     *
     * Typical usage:
     *
     * vlc v4l:/dev/video:input=2:norm=pal:size=192x144
     * --sout '#transcode{vcodec=theora,vb=300,acodec=vorb,ab=96}
     * :std{access=shout,mux=ogg,dst=localhost:8005}'
     *
     */

v4l refers to GNU/Linux Video4Linux and won't work for Windows users.

This comment precedes the genre option:

    /* To be listed properly as a public stream on the Yellow Pages of shoutcast/icecast
       the genres should match those used on the corresponding sites. Several examples
       are Alternative, Classical, Comedy, Country etc. */

This comment precedes the stream information options:

    /* The shout module only "transmits" data. It does not have direct access to
       "codec level" information. Stream information such as bitrate, samplerate,
       channel numbers and quality (in case of Ogg streaming) need to be set manually */

##### Options

-   **sout-shout-name \** : Name to give to this stream/channel on the shoutcast/icecast server *default value: "VLC media player - Live stream"*
-   **sout-shout-description \** : Description of the stream content or information about your channel *default value: "Live stream from VLC media player"*
-   **sout-shout-mp3 \** : You normally have to feed the shoutcast module with Ogg streams. It is also possible to stream MP3 instead, so you can forward MP3 streams to the shoutcast/icecast server *default value: disabled*
-   **sout-shout-genre \** : Genre of the content *default value: "Alternative"*
-   **sout-shout-url \** : URL with information about the stream or your channel *default value: "0*
-   **sout-shout-bitrate \** : Bitrate information of the transcoded stream *default value: ""*
-   **sout-shout-samplerate \** : Samplerate information of the transcoded stream *default value: ""*
-   **sout-shout-channels \** : Number of channels information of the transcoded stream *default value: ""*
-   **sout-shout-quality \** : Ogg Vorbis Quality information of the transcoded stream *default value: ""*
-   **sout-shout-public \** : Make the server publicly available on the 'Yellow Pages' (directory listing of streams) on the icecast/shoutcast website. Requires the bitrate information specified for shoutcast. Requires Ogg streaming for icecast *default value: disabled*

#### Services discovery

Module: shout

**Type**: Services discovery

**First VLC version**: 0.8.2

**Last VLC version**: 1.0.6

**Operating system(s)**: all

**Description**: Shoutcast services discovery module

**Shortcut(s)**: 0, 1

Three sub-modules had shortcuts of 0, 1 and 2.

##### Options

None (0 was deprecated with [\[acb5da732a27b6c7e8d6e05c2e183d4ae49a9ea9\]](https://git.videolan.org/?p=vlc/vlc-1.0.git;a=commitdiff;h=acb5da732a27b6c7e8d6e05c2e183d4ae49a9ea9)).

##### shout-winamp

This sub-module had the shortcut 0 with description "New winamp 5.2 shoutcast import". It is scheduled [to be removed](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=d3859f364921c6f4d48115da331ac3a44d7a6351) (currently in 4.0.0-dev) with the note:

    Removes the long unused Winamp/SHOUTcast directory stream filter for
    playlist handling, which was mostly useful together with the service
    discovery (modules/services_discovery/shout.c) which is not present
    anymore.

History:

-   [\[acb5da732a27b6c7e8d6e05c2e183d4ae49a9ea9\]](https://git.videolan.org/?p=vlc.git;a=commit;h=acb5da732a27b6c7e8d6e05c2e183d4ae49a9ea9) (introduction)
-   [modules/demux/playlist/shoutcast.c](https://git.videolan.org/?p=vlc/vlc-0.8.git;a=blob;f=modules/demux/playlist/shoutcast.c) (vlc/vlc-0.8.git)
-   [modules/demux/playlist/shoutcast.c](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=blob;f=modules/demux/playlist/shoutcast.c) (vlc/vlc-0.9.git)
-   [modules/demux/playlist/shoutcast.c](https://git.videolan.org/?p=vlc/vlc-1.0.git;a=blob;f=modules/demux/playlist/shoutcast.c) (vlc/vlc-1.0.git)
-   [modules/demux/playlist/shoutcast.c](https://git.videolan.org/?p=vlc/vlc-1.1.git;a=blob;f=modules/demux/playlist/shoutcast.c) (vlc/vlc-1.1.git)
-   [modules/demux/playlist/shoutcast.c](https://git.videolan.org/?p=vlc/vlc-2.0.git;a=blob;f=modules/demux/playlist/shoutcast.c) (vlc/vlc-2.0.git)
-   [modules/demux/playlist/shoutcast.c](https://git.videolan.org/?p=vlc/vlc-2.1.git;a=blob;f=modules/demux/playlist/shoutcast.c) (vlc/vlc-2.1.git)
-   [modules/demux/playlist/shoutcast.c](https://git.videolan.org/?p=vlc/vlc-2.2.git;a=blob;f=modules/demux/playlist/shoutcast.c) (vlc/vlc-2.2.git)
-   [modules/demux/playlist/shoutcast.c](https://git.videolan.org/?p=vlc/vlc-3.0.git;a=blob;f=modules/demux/playlist/shoutcast.c) (vlc/vlc-3.0.git)
-   [\[d3859f364921c6f4d48115da331ac3a44d7a6351\]](https://git.videolan.org/?p=vlc.git;a=commit;h=d3859f364921c6f4d48115da331ac3a44d7a6351) (removal)

#### Source code

-   [modules/access_output/shout.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access_output/shout.c)
-   [modules/services_discovery/shout.c](https://git.videolan.org/?p=vlc/vlc-1.0.git;a=blob;f=modules/services_discovery/shout.c) (vlc/vlc-1.0.git)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### skins2 {#modules-skins2}

VLC media player supports skins (sometimes also called themes) through its *skins2* interface module. To get new skins go to the [Skins website](http://www.videolan.org/vlc/skins.php).

*The steps mentioned here apply to VLC 0.9 and upward.*

**Skins are not available on macOS.**

*If you do have problems with VLC after applying a skin, a reinstall is NOT necessary. See #How do I fix VLC when it does not anymore show up properly.*

#### How to switch to the Skins interface

In order that you can use skins you have to change from VLC's native interface to the skinnable one. You can do that by opening the preferences and choosing *Use custom skin* in the section *Look and feel*. Then click the *Save* button and **restart** VLC. It should then show up in the default skin.


or use command-line to start VLC with skins interface

    vlc -I skins2

#### How to get new skins and where to save them

You can download a variety of skins from the [Skins website](http://www.videolan.org/vlc/skins.php). They usually should come as files with the extension "VLT". Although your browser or operating system might identify them as archive files, don't unpack them. Rather put them as they are into the skins folder of VLC.

On **Windows** this folder is located in the installation directory of VLC,

e.g. ***C:\\Program Files\\VideoLAN\\VLC\\skins***.

On **Linux/Unix** it is located in

***\~/.local/share/vlc/skins2***.

If you downloaded the skin pack just unpack it to the folders mentioned above.

***Warning: Not all of the skins available for download are fully functional.***

#### How to change to another skin

To change to a downloaded skin, **right-click** anywhere on the skin's background and choose *Interface*. Then chose either *Select skin* to change to a skin that is located in the default skin folder of VLC or *Open skin* to open a skin file that is located elsewhere.

#### How to switch back to VLC's default interface

When you open VLC and the skin you chose appears, right-click somewhere on the skins background and then choose *Interface* and *Preferences* (also accessible by pressing Ctrl+P). In the preferences dialog change the *Interface type* to *Native*. Then click save and restart VLC. It should show up in its default interface.

#### How do I fix VLC when it does not anymore show up properly

If you chose a broken skin it might happen that VLC does not anymore show up properly and that you cannot anymore access the settings as mentioned above.

Then you have to switch back to the default interface.

On Windows you can just open your Start menu and open

*All programs \> VideoLAN \> Quick Settings \> Interface \> Set Main Interface to Qt (default)*

On any other system, or when the start menu entry is missing, run VLC with the following command line:

    vlc -I qt

Now that VLC has been started with its native interface you can open the preferences (Ctrl+P) and change the active skin file. Chose the default skin or a skin you know that works. Then again set the skin interface to be the default one and restart VLC.

#### Are there skins with a full screen controller?

Full screen controllers in skins are supported since VLC 1.1. But apart from the default skin coming with VLC not many other skins have this feature.

#### How to create your own skin

There exists a program that enables you to create skins without any deep knowledge how skins are made up exactly. It is the [VLC Skin Editor](https://www.videolan.org/vlc/skineditor.html)

If you'd rather want to explore all the possibilities of the skin system and get to know how skins are made up and how to create them in detail, check out the [Skins2 documentation](https://www.videolan.org/vlc/skins2-create.html).

If you have any problems while creating your skin, please turn to the [skins forum](http://forum.videolan.org/viewforum.php?f=15).

#### See also

-   Skins2 Contest (contest over)

### smem {#modules-smem}

See also: Stream to memory (smem) tutorial_tutorial/ "Stream to memory (smem) tutorial")

Module: smem

**Type**: Stream output

**First VLC version**: 1.1.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Stream output to memory buffer

**Shortcut(s)**: 0

There is a note in the source code:

    /*
     * How to use it
     *
     *
     * You should use this module in combination with the transcode module, to get
     * raw datas from it. This module does not make any conversion at all, so you
     * need to use the transcode module for this purpose.
     *
     * For example, you can use smem as it :
     * --sout="#transcode{vcodec=RV24,acodec=s16l}:smem{smem-options}"
     *
     * Into each lock function (audio and video), you will have all the information
     * you need to allocate a buffer, so that this module will copy data in it.
     *
     * the video-data and audio-data pointers will be passed to lock/unlock function
     *
     **/

#### Options

-   **sout-smem-video-prerender-callback \** : Address of the video prerender callback function. This function will set the buffer where render will be done. *default value: "0"*
-   **sout-smem-audio-prerender-callback \** : Address of the audio prerender callback function. This function will set the buffer where render will be done. *default value: "0"*
-   **sout-smem-video-postrender-callback \** : Address of the video postrender callback function. This function will be called when the render is into the buffer. *default value: "0"*
-   **sout-smem-audio-postrender-callback \** : Address of the audio postrender callback function. This function will be called when the render is into the buffer. *default value: "0"*
-   **sout-smem-video-data \** : Data for the video callback function. *default value: "0"*
-   **sout-smem-audio-data \** : Data for the audio callback function. *default value: "0"*
-   **sout-smem-time-sync \** : Time Synchronisation option for output. If true, stream will render as usual, else it will be rendered as fast as possible. *default value: enabled*

#### Source code

-   [modules/stream_out/smem.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/stream_out/smem.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### speex {#modules-speex}

This module consists of a decoder, packetiser submodule and encoder submodule. Only the encoder has any options.

#### Demux

Module: speex

**Type**: Access demux

**First VLC version**: 0.7.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Speex audio decoder

**Shortcut(s)**: (none)

##### Options

None.

#### Packetizer

Module: speex

**Type**: Packetizer

**First VLC version**: 0.7.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Speex audio packetizer

**Shortcut(s)**: (none)

##### Options

None.

#### Mux

Module: speex

**Type**: Muxer

**First VLC version**: 0.7.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Speex audio encoder

**Shortcut(s)**: (none)

##### Options

-   **sout-speex-mode \ {0,1,2}** : Enforce the mode of the encoder: 0 - Narrow-band (8kHz), 1 - Wide-band (16kHz), 2 - Ultra-wideband (32kHz) *default value: 0*
-   **sout-speex-complexity \** : Enforce the complexity of the encoder *default value: 3*
-   **sout-speex-cbr \** : Enforce a constant bitrate encoding (CBR) instead of default variable bitrate encoding (VBR) *default value: disabled*
-   **sout-speex-quality \** : Enforce a quality between 0 (low) and 10 (high) *default value: 8.0*
-   **sout-speex-max-bitrate \** : Enforce the maximal VBR bitrate *default value: 0*
-   **sout-speex-vad \** : Enable [voice activity detection](http://en.wikipedia.org/wiki/voice_activity_detection) (VAD). It is automatically activated in VBR mode *default value: enabled*
-   **sout-speex-dtx \** : Enable [discontinuous transmission](http://en.wikipedia.org/wiki/discontinuous_transmission) (DTX) *default value: disabled*

#### Source code

-   [modules/codec/speex.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/speex.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### standard {#modules-standard}

Module: stream_out_standard

**Type**: Stream output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Standard stream output module

**Shortcut(s)**: 0, 1

The option 0 was deprecated in 2.1.0. The option 1 was deprecated in 3.0.0.

-   **sout-standard-access \** : Output method to use for the stream. *default value: ""*
-   **sout-standard-mux \** : Muxer to use for the stream. *default value: ""*
-   **sout-standard-dst \** : Destination (URL) to use for the stream. Overrides path and bind parameters. *default value: ""*
-   **sout-standard-bind \** : Address to bind to (helper setting for dst) address:port to bind vlc to listening incoming streams. Helper setting for dst, 0. dst-parameter overrides this. *default value: ""*
-   **sout-standard-path \** : Filename for stream. Helper setting for dst, 0. dst-parameter overrides this. *default value: ""*
-   **sout-standard-sap \** : Announce this session with SAP. *default value: disabled*
-   **sout-standard-name \** : This is the name of the session that will be announced in the SDP (Session Descriptor). *default value: ""*
-   **sout-standard-description \** : This allows you to give a short description with details about the stream, that will be announced in the SDP (Session Descriptor). *default value: ""*
-   **sout-standard-url \** : This allows you to give a URL with more details about the stream (often the website of the streaming organization), that will be announced in the SDP (Session Descriptor). *default value: ""*
-   **sout-standard-email \** : This allows you to give a contact mail address for the stream, that will be announced in the SDP (Session Descriptor). *default value: ""*

#### Source code

-   [modules/stream_out/standard.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/stream_out/standard.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### std {#modules-std}

Module: stream_out_standard

**Type**: Stream output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Standard stream output module

**Shortcut(s)**: 0, 1

The option 0 was deprecated in 2.1.0. The option 1 was deprecated in 3.0.0.

-   **sout-standard-access \** : Output method to use for the stream. *default value: ""*
-   **sout-standard-mux \** : Muxer to use for the stream. *default value: ""*
-   **sout-standard-dst \** : Destination (URL) to use for the stream. Overrides path and bind parameters. *default value: ""*
-   **sout-standard-bind \** : Address to bind to (helper setting for dst) address:port to bind vlc to listening incoming streams. Helper setting for dst, 0. dst-parameter overrides this. *default value: ""*
-   **sout-standard-path \** : Filename for stream. Helper setting for dst, 0. dst-parameter overrides this. *default value: ""*
-   **sout-standard-sap \** : Announce this session with SAP. *default value: disabled*
-   **sout-standard-name \** : This is the name of the session that will be announced in the SDP (Session Descriptor). *default value: ""*
-   **sout-standard-description \** : This allows you to give a short description with details about the stream, that will be announced in the SDP (Session Descriptor). *default value: ""*
-   **sout-standard-url \** : This allows you to give a URL with more details about the stream (often the website of the streaming organization), that will be announced in the SDP (Session Descriptor). *default value: ""*
-   **sout-standard-email \** : This allows you to give a contact mail address for the stream, that will be announced in the SDP (Session Descriptor). *default value: ""*

#### Source code

-   [modules/stream_out/standard.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/stream_out/standard.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### subsdelay {#modules-subsdelay}

Module: subsdelay

**Type**: Video sub-filter

**First VLC version**: 1.2.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Change subtitles delay

**Shortcut(s)**: -

The subsdelay filter can help slow readers to keep up with the subtitles.
It extends the subtitles duration without changing their original appearance time, so the subtitles are piled up on the video. To help keep track of the appearance order, existing subtitles gets more transparent as new subtitles arrive.

The subtitles duration factor is configurable through the synchronization dialog. Other options can be set through the preferences (*show all settings* → Video → Subtitles/OSD → Subsdelay).

#### Options

-   **subsdelay-mode \ { 0, 1, 2 }** : Delay calculation mode *default value: 1*
-   **subsdelay-factor \** : The delay calculation parameter *default value: 2.0*
-   **subsdelay-overlap \** : Maximum number of subtitles allowed at the same time *default value: 3*
-   **subsdelay-min-alpha \** : Alpha value of the earliest subtitle, where 0 is fully transparent and 255 is fully opaque.
    Subtitles alpha is somewhere between fully opaque and this value according to the appearance order and the maximum overlapping allowed *default value: 70*

##### Overlap fix

These rules help fixing some "flickering" effects caused by the overlapping. They are applied after the initial delay is calculated in the following order:

-   **subsdelay-min-stops \** : Minimum time (in milliseconds) that a subtitle should stay after its predecessor has disappeared (subtitle delay will be extended to meet this requirement) *default value: 1000*
-   **subsdelay-min-stop-start \** : Minimum time (in milliseconds) between subtitle disappearance and a newer subtitle appearance (earlier subtitle delay will be extended to fill the gap) *default value: 1000*
-   **subsdelay-min-start-stop \** : Minimum time (in milliseconds) that a subtitle should stay after a newer subtitle has appeared (earlier subtitle delay will be shortened to avoid the overlap) *default value: 1000*

#### Examples

Example command line use **(VLC 1.2.0 and above)** :

    % vlc --sub-filter subsdelay --subsdelay-mode 1 --subsdelay-factor 2 --subsdelay-overlap 3

Multiply subtitles duration by 2, up to 3 subtitles can be overlapped at a given time.

    % vlc --sub-filter subsdelay --subsdelay-mode 0 --subsdelay-factor 0 --subsdelay-overlap 1 --subsdelay-min-stop-start 0

Don't change subtitles duration but fix any existing overlaps.

#### Source code

-   [modules/spu/subsdelay.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/spu/subsdelay.c)

#### Appendix

For option --subsdelay-mode:

0
:   0
    Absolute delay - add an absolute delay to each subtitle.
    In this mode the factor represents seconds

1
:   0
    Relative to source delay - multiply subtitles delay.

2
:   0
    Relative to source content - determine subtitles delay from its content.
    The delay calculation is based on the number and length of the words in the subtitle.
    This mode could only work for plain subtitles sources (like SubRip, MicroDVD, etc), for other formats the "relative to source delay" mode is used instead

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### subtitle {#modules-subtitle}

Module: subtitle

**Type**: Access demux

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Text subtitle parser

**Shortcut(s)**: 0

Option 0 was removed in [\[204eb2a0ea3bca9d58002adab5ad937aa2e1ac7c\]](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=204eb2a0ea3bca9d58002adab5ad937aa2e1ac7c) and option 1 was removed in [\[28d124dd6567e120ee730f8a02395089e65ba79f\]](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=28d124dd6567e120ee730f8a02395089e65ba79f). They are now in libVLC.

#### Options

-   **sub-type \ { auto, microdvd, subrip, subviewer, ssa1, ssa2-4, ass, vplayer, sami, dvdsubtitle, mpl2, aqt, pjs, mpsub, jacosub, psb, realtext, dks, subviewer1, sbv }** : Force the subtiles format. Selecting "auto" means autodetection and should always work *default value: auto*
-   **sub-description \** : Override the default track description *default value: NULL*

#### Source code

-   [modules/demux/subtitle.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/subtitle.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### svg {#modules-svg}

Decoding and encoding (text rendering) are handled by separate modules. Both modules have the same shortcut, 0, though [modules/MODULES_LIST](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/MODULES_LIST) calls them 1 and 2.

#### Decoder

Module: svg

**Type**: Input

**First VLC version**: 2.2.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: SVG video decoder making use of librsvg2

**Shortcut(s)**: -

##### Options

-   **svg-width \** : Specify the width to decode the image to *default value: -1*
-   **svg-height \** : Specify the height to decode the image to *default value: -1*
-   **svg-scale \** : Scale factor to apply to image *default value: -1.0*

#### Encoder

Module: svg

**Type**: Input

**First VLC version**: 0.8.0

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: Put SVG on the video

**Shortcut(s)**: -

##### Options

-   **svg-template-file \** : Location of a file holding a SVG template for automatic string conversion *default value: ""*

#### Source code

-   [modules/codec/svg.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/svg.c) (decoder)
-   [modules/text_renderer/svg.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/text_renderer/svg.c) (encoder)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### switcher {#modules-switcher}

Module: switcher

**Type**: Stream output

**First VLC version**: -

**Last VLC version**: 2.0.9

**Operating system(s)**: all

**Description**: MPEG-2 video switcher stream output

**Shortcut(s)**: 0

This module used the avcodec library.

#### Options

-   **sout-switcher-files \** : Full paths of the files separated by colons *default value: ""*
-   **sout-switcher-sizes \** : List of sizes separated by colons (720x576:480x576) *default value: ""*
-   **sout-switcher-aspect-ratio \** : Aspect ratio (4:3, 16:9) *default value: 4:3*
-   **sout-switcher-port \** : UDP port to listen to for commands *default value: 5001*
-   **sout-switcher-command \** : Initial command to execute *default value: 0*
-   **sout-switcher-gop \** : Number of P-frames between two I-frames *default value: 8*
-   **sout-switcher-qscale \** : Fixed quantizer scale to use *default value: 5*
-   **sout-switcher-mute-audio \** : Mute audio when command is not 0 *default value: enabled*

#### Source code

-   [modules/stream_out/switcher.c](https://git.videolan.org/?p=vlc/vlc-2.0.git;a=blob;f=modules/stream_out/switcher.c) (vlc/vlc-2.0.git)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### telnet {#modules-telnet}

See also: Documentation:Modules/ncurses and Documentation:Modules/http intf

Module: telnet

**Type**: Control interface

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Control VLC via a telnet connection

**Shortcut(s)**: -

The telnet module communicates with VLC over a network connection using the [telnet](http://en.wikipedia.org/wiki/telnet) protocol. The original module was provided until 1.1.0, when it was re-written in [Lua](http://en.wikipedia.org/wiki/Lua_(programming_language) "wikipedia:Lua (programming language)"). The old module was renamed to oldtelnet and removed in 2.0.0.

Telnet should not be used for sensitive applications.

To find module information on the command-line for VLC 2.0.0 and above, use 0 and look for the *Lua Telnet* section.

Options as of 3.0.6 are listed below:

-   **telnet-host \** : This is the host on which the interface will listen. It defaults to all network interfaces (0.0.0.0). If you want this interface to be available only on the local machine, enter "127.0.0.1"
-   **telnet-port \** : This is the TCP port on which this interface will listen *default value: 4212*
-   **telnet-password \** : A single password restricts access to this interface
-   **lua-sd \** : ?

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### telx {#modules-telx}

Module: telx

**Type**: Access demux

**First VLC version**: 0.8.6b

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Teletext subtitles decoder

**Shortcut(s)**: (none)

#### Options

-   **telx-override-page \** : Override the indicated page, try this if your subtitles don't appear (-1 = autodetect from TS, 0 = autodetect from teletext, \>0 = actual page number, usually 888 or 889) *default value: -1*
-   **telx-ignore-subtitle-flag \** : Ignore the subtitle flag, try this if your subtitles don't appear *default value: disabled*
-   **telx-french-workaround \** : Some French channels do not flag their subtitling pages correctly due to a historical interpretation mistake. Try using this wrong interpretation if your subtitles don't appear *default value: disabled*

#### Source code

-   [modules/codec/telx.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/telx.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### time {#modules-time}

**This page is obsolete and kept only for historical interest.** It may document features that are obsolete, superseded, or irrelevant. Do not rely on the information here being up-to-date.
Additional information: **This filter has been merged with the marq filter in version 0.9.0.**

Module: time

**Type**: Video sub-filter

**First VLC version**: 0.8.0

**Last VLC version**: 0.8.6

**Operating system(s)**: all

**Description**: Overlays date and time on the video

**Shortcut(s)**: 0

Allows overlaying date and time information on the video.

#### Options

The option for the time picture subfilter in version 0.8.6 are the following:

-   **time-format \** : Time format string (%Y%m%d %H%M%S) *default value: "%Y-%m-%d %H:%M:%S"*
-   **time-x \** : X offset *default value: -1*
-   **time-y \** : Y offset *default value: 0*
-   **time-position \ { 0, 1, 2, 4, 8, 5, 6, 9, 10 }** : Position *default value: 9*
-   **time-opacity \** : Opacity *default value: 255*
-   **time-color \ { -268435456, 0, 8421504, 12632256, 16777215, 8388608, 16711680, 16711935, 16776960, 8421376, 32768, 32896, 65280, 8388736, 128, 255, 65535 }** : Colour^(**key**)^ *default value: 16777215*
-   **time-size \** : Font size, pixels *default value: -1*

#### Usage

There are two ways to use the time module: over screen output or display; and over transcoded output.

##### Screen output or display

To overlay the current time over vlc screen output or display, use the --time-? options (where ? means "format," "x", "y" etc; i.e. --time-format).

In this example, the time will be displayed in white on the lower right hand corner of the viewable output of a transcoded stream and sent to a multicast IP address with the associated SAP announce.

    % vlc input_stream --sub-filter=time --time-format %Y-%m-%d,%H:%M:%S --time-position 9 --time-color 16777215 --time-size 12 --sout "#transcode{venc=ffmpeg,vcodec=mp4v}:duplicate{dst=display,dst=rtp{mux=ts,dst=239.255.12.42,sdp=sap,name="TestStream"}}"

In this example, the time will be displayed as 2007-6-19,10:09:33. In addition, the time will only be displayed on the visual display of the input_stream. It will not be part of the transcoded output.

##### Transcoded output

To overlay the current time over the transcoded output, enable the transcode module subpicture filter or sfilter option.

In this example, the time will be displayed in white on the lower right only in the transcoded output.

    % vlc input_stream --time-format %Y-%m-%d,%H:%M:%S --time-position 9 --time-color 16777215 --time-size 12 --sout "#transcode{venc=ffmpeg,vcodec=mp4v,sfilter=time}:duplicate{dst=display,dst=rtp{mux=ts,dst=239.255.12.42,sdp=sap,name="TestStream"}}"

Note that this is accomplished by removing the --sub-filter=time command line option and adding the sfilter transcode module option. If the --sub-filter=time is included vlc will overlay the time over the overlay transcode time, essentially overlapping it.

Also note that the --time-? command line options are "global;" i.e., they affect the way the time overlays both the display and the transcoded output.

#### Source code

-   [modules/video_filter/time.c](https://git.videolan.org/?p=vlc/vlc-0.8.git;a=blob;f=modules/video_filter/time.c) (vlc/vlc-0.8.git)

#### Appendix

-   \^ --time-position
-   \^ --time-color

&nbsp;

-   **Integer**: Alignment Comment
-   **0**: Center
-   **1**: Left
-   **2**: Right
-   **4**: Top
-   **8**: Bottom
-   **5**: Top-Left 4 + 1
-   **6**: Top-Right 4 + 2
-   **9**: Bottom-Left 8 + 1
-   **10**: Bottom-Right 8 + 2
-   **3**: n/a contradictory
-   **7**: n/a contradictory

&nbsp;

-   **Sample**: Integer code Colour
-   **0**: Default
-   **0**: Black
-   **0**: Gray
-   **0**: Silver
-   **0**: White
-   **0**: Maroon
-   **0**: Red
-   **0**: Fuchsia
-   **0**: Yellow
-   **0**: Olive
-   **0**: Green
-   **0**: Teal
-   **0**: Lime
-   **0**: Purple
-   **0**: Navy
-   **0**: Blue
-   **0**: Aqua

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### timeshift {#modules-timeshift}

Module: timeshift

**Type**: Access filter

**First VLC version**: 0.8.2

**Last VLC version**: 0.9.9

**Operating system(s)**: all

**Description**: enable timeshifting on live streams

**Shortcut(s)**: -

This access filter will enable timeshifting on live streams. The user will thus be able to pause the stream. Buffered data will be stored in memory for short periods and on the hard drive afterwards.

**\*\* Warning : the following documentation is deprecated \*\***

It is now in the VLC core.

#### Options

-   **timeshift-granularity \** : Size of temporary files in MB *default value: 50*

&nbsp;

-   **timeshift-dir \** : Directory where temporary files will be stored.

&nbsp;

-   **timeshift-force** : Force use of the timeshift module even if the underlying access claims that it can pause *default value: disabled*

#### Example

    % vlc --access-filter timeshift

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### transcode {#modules-transcode}

Module: stream_out_transcode

**Type**: Stream output

**First VLC version**: 0.6.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Transcode content on the fly

**Shortcut(s)**: -

The only shortcut to this module is 0.

#### Options

Note: Since supported codecs are dynamically assigned by the running program, 0, 1 and 2 have been left blank.

Looking at the source code for 4.0.0-dev it seems no checks are directly performed limiting 0 beyond 1.

As of 2.2.0 0 accepts fps as rationals e.g. 1.

Deprecated options:

-   0 (since 2.2.0), 1 seems to be equivalent
-   0 (since 2.2.0)

##### Video

-   **sout-transcode-venc \** : This is the video encoder module that will be used (and its associated options)
-   **sout-transcode-vcodec \** : This is the video codec that will be used *default value: NULL*
-   **sout-transcode-vb \** : Target bitrate of the transcoded video stream *default value: 0*
-   **sout-transcode-scale \** : Scale factor to apply to the video while transcoding (eg: 0.25) *default value: 0*
-   **sout-transcode-fps \** : Target output frame rate for the video stream *default value: NULL*
-   **sout-transcode-deinterlace \** : Deinterlace the video before encoding *default value: disabled*
-   **sout-transcode-deinterlace-module \ {deinterlace,ffmpeg-deinterlace}** : Specify the deinterlace module to use *default value: deinterlace*
-   **sout-transcode-width \** : Output video width *default value: 0*
-   **sout-transcode-height \** : Output video height *default value: 0*
-   **sout-transcode-maxwidth \** : Maximum output video width *default value: 0*
-   **sout-transcode-maxheight \** : Maximum output video height *default value: 0*
-   **sout-transcode-vfilter \** : Video filters will be applied to the video streams (after overlays are applied). You can enter a colon-separated list of filters

##### Audio

-   **sout-transcode-aenc \** : This is the audio encoder module that will be used (and its associated options)
-   **sout-transcode-acodec \** : This is the audio codec that will be used *default value: NULL*
-   **sout-transcode-ab \** : Target bitrate of the transcoded audio stream *default value: 96*
-   **sout-transcode-alang \** : This is the language of the audio stream *default value: NULL*
-   **sout-transcode-channels \** : Number of audio channels in the transcoded streams *default value: 0*
-   **sout-transcode-samplerate \** : Sample rate of the transcoded audio stream (11250, 22500, 44100 or 48000) *default value: 0*
-   **sout-transcode-afilter \** : Audio filters will be applied to the audio streams (after conversion filters are applied). You can enter a colon-separated list of filters

##### Overlays/Subtitles

-   **sout-transcode-senc \** : This is the subtitle encoder module that will be used (and its associated options)
-   **sout-transcode-scodec \** : This is the subtitle codec that will be used *default value: NULL*
-   **sout-transcode-soverlay \** : This is the subtitle codec that will be used *default value: disabled*
-   **sout-transcode-sfilter \** : This allows you to add overlays (also known as "subpictures") on the transcoded video stream. The subpictures produced by the filters will be overlayed directly onto the video. You can specify a colon-separated list of subpicture modules

##### Miscellaneous

-   **sout-transcode-threads \** : Number of threads used for the transcoding *default value: 0*
-   **sout-transcode-pool-size \** : Defines how many pictures we allow to be in pool between decoder/encoder threads when threads \> 0 *default value: 10*
-   **sout-transcode-high-priority \** : Runs the optional encoder thread at the OUTPUT priority instead of VIDEO *default value: disabled*

#### Source code

-   [modules/stream_out/transcode/transcode.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/stream_out/transcode/transcode.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### transform {#modules-transform}

Module: transform

**Type**: Video output filter

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Rotate or flip the video

**Shortcut(s)**: 0

#### Options

-   **transform-type \ { "90", "180", "270", "hflip", "vflip", "transpose", "antitranspose" }** : Transformation type *default value: "90"*

#### Examples

    $ vlc --video-filter='transform{type="vflip"}' somevideo.avi

#### Source code

-   [modules/video_filter/transform.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/transform.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### transrate {#modules-transrate}

Module: transrate

**Type**: Stream output

**First VLC version**: 0.7.0

**Last VLC version**: 1.0.2

**Operating system(s)**: all

**Description**: MPEG-2 video transrating stream output

**Shortcut(s)**: -

A simplified history: this module was [introduced](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=c37e84c1f930e22a99d03e76a6a7f2a4be3ed420) in VLC 0.7.0 and [removed completely](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=29d2ca92fc94a9407ed829539fe3403c6499970e) prior to VLC 1.1.0.

#### Options

None.

#### Source code

-   [modules/stream_out/transrate/transrate.c](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=blob;f=modules/stream_out/transrate/transrate.c) (vlc/vlc-0.9.git)
-   [modules/stream_out/transrate](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=tree;f=modules/stream_out/transrate;hb=HEAD) (vlc/vlc-0.9.git)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### udp {#modules-udp}

#### Access

Module: udp

**Type**: Access

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: UDP input

**Shortcut(s)**: 0, 1, 2, 3

The options 0 and 1 were deprecated in 2.0.0 and 3.0.0. 2 was added in 3.0.0.

-   **udp-timeout \** : UDP Source timeout (sec) *default value: -1*

#### Access output

Module: udp

**Type**: Access output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: UDP stream output

**Shortcut(s)**: 0

-   **sout-udp-caching \** : Default caching value for outbound UDP streams. This value should be set in milliseconds *default value: 0{.variable}1*
-   **sout-udp-group \** : Packets can be sent one by one at the right time or by groups. You can choose the number of packets that will be sent at a time. It helps reducing the scheduling load on heavily-loaded systems *default value: 1*

#### Source code

-   [modules/access/udp.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/udp.c)
-   [modules/access_output/udp.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access_output/udp.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### upnp {#modules-upnp}

Prior to VLC 2.0.0 the [UPnP](http://en.wikipedia.org/wiki/UPnP) module had 2 files: upnp_cc (for [CyberLink](http://en.wikipedia.org/wiki/CyberLink)) and upnp_intel (for [Intel](http://en.wikipedia.org/wiki/Intel_Corporation)).
The upnp_cc file was [removed](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=63751e5aef7dc2ef5098df0df8bdca07849d8fd5) and the upnp_intel file was [renamed](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=15e31aa8a7a30df086bb31422b750dcbd632dfae) to upnp.

#### upnp.cpp

Module: upnp

**Type**: Services discovery

**First VLC version**: 0.8.4

**Last VLC version**: -

**Operating system(s)**: all

**Description**: [Universal Plug'n'Play](http://en.wikipedia.org/wiki/Universal_Plug%27n%27Play)

**Shortcut(s)**: (none)

When VLC is compiled with UPNP support, you will still need^\[May\ no\ longer\ be\ necessary?\]^ to enable UPNP service discovery:

-   either on command line via \$ vlc --services-discovery upnp_intel
-   or in the playlist menu: File/Service discovery/UPNP

Then discovered UPNP servers will be listed on the playlist.

##### Options

Note the spelling difference: it is option satip-channe**l**ist and satip-channe**ll**ist-url.

-   **satip-channelist \ { "Auto", "ASTRA_19_2E", "ASTRA_28_2E", "ASTRA_23_5E", "MasterList", "ServerList", "CustomList" }** : Custom SAT\>IP channel list URL *default value: "auto"*
-   **satip-channellist-url \** : Custom SAT\>IP channel list URL *default value: NULL*

##### upnp

###### Options

None.

##### upnp_renderer

Module: upnp_renderer

**Type**: Renderer discovery

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: UPnP Renderer Discovery

**Shortcut(s)**: 0

###### Options

None.

#### upnp_cc.cpp

Module: UPnP

**Type**: Services discovery

**First VLC version**: -

**Last VLC version**: 1.1.?

**Operating system(s)**: all

**Description**: Universal Plug'n'Play

**Shortcut(s)**: (none)

##### Options

None.

#### Source code

Current:

-   [modules/services_discovery/upnp.cpp](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/services_discovery/upnp.cpp)

Former:

-   [modules/services_discovery/upnp_cc.cpp](https://git.videolan.org/?p=vlc/vlc-0.8.git;a=blob;f=modules/services_discovery/upnp_cc.cpp) (vlc/vlc-0.8.git)
-   [modules/services_discovery/upnp_cc.cpp](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=blob;f=modules/services_discovery/upnp_cc.cpp) (vlc/vlc-0.9.git)
-   [modules/services_discovery/upnp_cc.cpp](https://git.videolan.org/?p=vlc/vlc-1.0.git;a=blob;f=modules/services_discovery/upnp_cc.cpp) (vlc/vlc-1.0.git)
-   [modules/services_discovery/upnp_cc.cpp](https://git.videolan.org/?p=vlc/vlc-1.1.git;a=blob;f=modules/services_discovery/upnp_cc.cpp) (vlc/vlc-1.1.git)
-   [modules/services_discovery/upnp_intel.cpp](https://git.videolan.org/?p=vlc/vlc-0.8.git;a=blob;f=modules/services_discovery/upnp_intel.cpp) (vlc/vlc-0.8.git)
-   [modules/services_discovery/upnp_intel.cpp](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=blob;f=modules/services_discovery/upnp_intel.cpp) (vlc/vlc-0.9.git)
-   [modules/services_discovery/upnp_intel.cpp](https://git.videolan.org/?p=vlc/vlc-1.0.git;a=blob;f=modules/services_discovery/upnp_intel.cpp) (vlc/vlc-1.0.git)
-   [modules/services_discovery/upnp_intel.cpp](https://git.videolan.org/?p=vlc/vlc-1.1.git;a=blob;f=modules/services_discovery/upnp_intel.cpp) (vlc/vlc-1.1.git)

#### Appendix

**For the option 0:**

-   **Option name**: Auto ASTRA_19_2E ASTRA_28_2E ASTRA_23_5E MasterList ServerList CustomList
-   **Meaning**: Auto Astra 19.2°E Astra 28.2°E Astra 23.5°E SAT\>IP Main List Device List Custom List

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### upnp cc {#modules-upnp-cc}

Prior to VLC 2.0.0 the [UPnP](http://en.wikipedia.org/wiki/UPnP) module had 2 files: upnp_cc (for [CyberLink](http://en.wikipedia.org/wiki/CyberLink)) and upnp_intel (for [Intel](http://en.wikipedia.org/wiki/Intel_Corporation)).
The upnp_cc file was [removed](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=63751e5aef7dc2ef5098df0df8bdca07849d8fd5) and the upnp_intel file was [renamed](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=15e31aa8a7a30df086bb31422b750dcbd632dfae) to upnp.

#### upnp.cpp

Module: upnp

**Type**: Services discovery

**First VLC version**: 0.8.4

**Last VLC version**: -

**Operating system(s)**: all

**Description**: [Universal Plug'n'Play](http://en.wikipedia.org/wiki/Universal_Plug%27n%27Play)

**Shortcut(s)**: (none)

When VLC is compiled with UPNP support, you will still need^\[May\ no\ longer\ be\ necessary?\]^ to enable UPNP service discovery:

-   either on command line via \$ vlc --services-discovery upnp_intel
-   or in the playlist menu: File/Service discovery/UPNP

Then discovered UPNP servers will be listed on the playlist.

##### Options

Note the spelling difference: it is option satip-channe**l**ist and satip-channe**ll**ist-url.

-   **satip-channelist \ { "Auto", "ASTRA_19_2E", "ASTRA_28_2E", "ASTRA_23_5E", "MasterList", "ServerList", "CustomList" }** : Custom SAT\>IP channel list URL *default value: "auto"*
-   **satip-channellist-url \** : Custom SAT\>IP channel list URL *default value: NULL*

##### upnp

###### Options

None.

##### upnp_renderer

Module: upnp_renderer

**Type**: Renderer discovery

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: UPnP Renderer Discovery

**Shortcut(s)**: 0

###### Options

None.

#### upnp_cc.cpp

Module: UPnP

**Type**: Services discovery

**First VLC version**: -

**Last VLC version**: 1.1.?

**Operating system(s)**: all

**Description**: Universal Plug'n'Play

**Shortcut(s)**: (none)

##### Options

None.

#### Source code

Current:

-   [modules/services_discovery/upnp.cpp](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/services_discovery/upnp.cpp)

Former:

-   [modules/services_discovery/upnp_cc.cpp](https://git.videolan.org/?p=vlc/vlc-0.8.git;a=blob;f=modules/services_discovery/upnp_cc.cpp) (vlc/vlc-0.8.git)
-   [modules/services_discovery/upnp_cc.cpp](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=blob;f=modules/services_discovery/upnp_cc.cpp) (vlc/vlc-0.9.git)
-   [modules/services_discovery/upnp_cc.cpp](https://git.videolan.org/?p=vlc/vlc-1.0.git;a=blob;f=modules/services_discovery/upnp_cc.cpp) (vlc/vlc-1.0.git)
-   [modules/services_discovery/upnp_cc.cpp](https://git.videolan.org/?p=vlc/vlc-1.1.git;a=blob;f=modules/services_discovery/upnp_cc.cpp) (vlc/vlc-1.1.git)
-   [modules/services_discovery/upnp_intel.cpp](https://git.videolan.org/?p=vlc/vlc-0.8.git;a=blob;f=modules/services_discovery/upnp_intel.cpp) (vlc/vlc-0.8.git)
-   [modules/services_discovery/upnp_intel.cpp](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=blob;f=modules/services_discovery/upnp_intel.cpp) (vlc/vlc-0.9.git)
-   [modules/services_discovery/upnp_intel.cpp](https://git.videolan.org/?p=vlc/vlc-1.0.git;a=blob;f=modules/services_discovery/upnp_intel.cpp) (vlc/vlc-1.0.git)
-   [modules/services_discovery/upnp_intel.cpp](https://git.videolan.org/?p=vlc/vlc-1.1.git;a=blob;f=modules/services_discovery/upnp_intel.cpp) (vlc/vlc-1.1.git)

#### Appendix

**For the option 0:**

-   **Option name**: Auto ASTRA_19_2E ASTRA_28_2E ASTRA_23_5E MasterList ServerList CustomList
-   **Meaning**: Auto Astra 19.2°E Astra 28.2°E Astra 23.5°E SAT\>IP Main List Device List Custom List

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### upnp intel {#modules-upnp-intel}

Prior to VLC 2.0.0 the [UPnP](http://en.wikipedia.org/wiki/UPnP) module had 2 files: upnp_cc (for [CyberLink](http://en.wikipedia.org/wiki/CyberLink)) and upnp_intel (for [Intel](http://en.wikipedia.org/wiki/Intel_Corporation)).
The upnp_cc file was [removed](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=63751e5aef7dc2ef5098df0df8bdca07849d8fd5) and the upnp_intel file was [renamed](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=15e31aa8a7a30df086bb31422b750dcbd632dfae) to upnp.

#### upnp.cpp

Module: upnp

**Type**: Services discovery

**First VLC version**: 0.8.4

**Last VLC version**: -

**Operating system(s)**: all

**Description**: [Universal Plug'n'Play](http://en.wikipedia.org/wiki/Universal_Plug%27n%27Play)

**Shortcut(s)**: (none)

When VLC is compiled with UPNP support, you will still need^\[May\ no\ longer\ be\ necessary?\]^ to enable UPNP service discovery:

-   either on command line via \$ vlc --services-discovery upnp_intel
-   or in the playlist menu: File/Service discovery/UPNP

Then discovered UPNP servers will be listed on the playlist.

##### Options

Note the spelling difference: it is option satip-channe**l**ist and satip-channe**ll**ist-url.

-   **satip-channelist \ { "Auto", "ASTRA_19_2E", "ASTRA_28_2E", "ASTRA_23_5E", "MasterList", "ServerList", "CustomList" }** : Custom SAT\>IP channel list URL *default value: "auto"*
-   **satip-channellist-url \** : Custom SAT\>IP channel list URL *default value: NULL*

##### upnp

###### Options

None.

##### upnp_renderer

Module: upnp_renderer

**Type**: Renderer discovery

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: UPnP Renderer Discovery

**Shortcut(s)**: 0

###### Options

None.

#### upnp_cc.cpp

Module: UPnP

**Type**: Services discovery

**First VLC version**: -

**Last VLC version**: 1.1.?

**Operating system(s)**: all

**Description**: Universal Plug'n'Play

**Shortcut(s)**: (none)

##### Options

None.

#### Source code

Current:

-   [modules/services_discovery/upnp.cpp](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/services_discovery/upnp.cpp)

Former:

-   [modules/services_discovery/upnp_cc.cpp](https://git.videolan.org/?p=vlc/vlc-0.8.git;a=blob;f=modules/services_discovery/upnp_cc.cpp) (vlc/vlc-0.8.git)
-   [modules/services_discovery/upnp_cc.cpp](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=blob;f=modules/services_discovery/upnp_cc.cpp) (vlc/vlc-0.9.git)
-   [modules/services_discovery/upnp_cc.cpp](https://git.videolan.org/?p=vlc/vlc-1.0.git;a=blob;f=modules/services_discovery/upnp_cc.cpp) (vlc/vlc-1.0.git)
-   [modules/services_discovery/upnp_cc.cpp](https://git.videolan.org/?p=vlc/vlc-1.1.git;a=blob;f=modules/services_discovery/upnp_cc.cpp) (vlc/vlc-1.1.git)
-   [modules/services_discovery/upnp_intel.cpp](https://git.videolan.org/?p=vlc/vlc-0.8.git;a=blob;f=modules/services_discovery/upnp_intel.cpp) (vlc/vlc-0.8.git)
-   [modules/services_discovery/upnp_intel.cpp](https://git.videolan.org/?p=vlc/vlc-0.9.git;a=blob;f=modules/services_discovery/upnp_intel.cpp) (vlc/vlc-0.9.git)
-   [modules/services_discovery/upnp_intel.cpp](https://git.videolan.org/?p=vlc/vlc-1.0.git;a=blob;f=modules/services_discovery/upnp_intel.cpp) (vlc/vlc-1.0.git)
-   [modules/services_discovery/upnp_intel.cpp](https://git.videolan.org/?p=vlc/vlc-1.1.git;a=blob;f=modules/services_discovery/upnp_intel.cpp) (vlc/vlc-1.1.git)

#### Appendix

**For the option 0:**

-   **Option name**: Auto ASTRA_19_2E ASTRA_28_2E ASTRA_23_5E MasterList ServerList CustomList
-   **Meaning**: Auto Astra 19.2°E Astra 28.2°E Astra 23.5°E SAT\>IP Main List Device List Custom List

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### v4l {#modules-v4l}

Module: v4l

**Type**: Access demux

**First VLC version**: -

**Last VLC version**: 1.1.13

**Operating system(s)**: Linux

**Description**: Video 4 Linux input

**Shortcut(s)**: -

#### Help Output

Each of these options can be used as commandline flags individually as shown below or as such: v4l:///dev/video:norm=secam:frequency=543250:size=640x480:channel=0:adev=/dev/dsp:audio=0 where you can see that /dev/video refers to vdev, and the rest line up with the flags below less the prefix --v4l-.

Video4Linux input (v4l)

         --v4l-caching     Caching value in ms
             Caching value for V4L captures. This value should be set in milliseconds.
         --v4l-vdev         Video device name
             Name of the video device to use. If you don't specify anything, no video device will be used.
         --v4l-adev         Audio device name
             Name of the audio device to use. If you don't specify anything, no audio device will be used.
         --v4l-chroma       Video input chroma format
             Force the Video4Linux video device to use a specific chroma format (eg. I420 (default), RV24, etc.)
         --v4l-fps           Framerate
             Framerate to capture, if applicable (-1 for autodetect).
         --v4l-samplerate  Samplerate
             Samplerate of the captured audio stream, in Hz (eg: 11025, 22050, 44100)
         --v4l-channel     Channel
             Channel of the card to use (Usually, 0 = tuner, 1 = composite, 2 = svideo).
         --v4l-tuner       Tuner
             Tuner to use, if there are several ones.
         --v4l-norm {3 (Automatic), 2 (SECAM), 0 (PAL), 1 (NTSC)}
                                    Norm
             Norm of the stream (Automatic, SECAM, PAL, or NTSC).
         --v4l-frequency   Frequency
             Frequency to capture (in kHz), if applicable.
         --v4l-audio       Audio Channel
             Audio Channel to use, if there are several audio inputs.
         --v4l-stereo, --no-v4l-stereo
                                    Stereo (default enabled)
             Capture the audio stream in stereo. (default enabled)
         --v4l-width       Width
             Width of the stream to capture (-1 for autodetect).
         --v4l-height      Height
             Height of the stream to capture (-1 for autodetect).
         --v4l-brightness  Brightness
             Brightness of the video input.
         --v4l-colour      Color
             Color of the video input.
         --v4l-hue         Hue
             Hue of the video input.
         --v4l-contrast    Contrast
             Contrast of the video input.
         --v4l-mjpeg, --no-v4l-mjpeg
                                    MJPEG (default disabled)
             Set this option if the capture device outputs MJPEG (default disabled)
         --v4l-decimation  Decimation
             Decimation level for MJPEG streams
         --v4l-quality     Quality
             Quality of the stream.

### v4l2 {#modules-v4l2}

Module: v4l2

**Type**: Access demux

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: Video for Linux 2 input

**Shortcut(s)**: -

#### Options

-   **v4l2-caching \** : Caching in ms

##### Video input

-   **v4l2-dev \** : Primary device name *default value: "/dev/video0"*
-   **v4l2-standard \** : Video standard *default value: 0*
-   **v4l2-chroma \** : Force use of a specific video chroma (Use MJPG here to use a webcam's MJPEG stream) *default value: ""*
-   **v4l2-input \** : Card input to use for video *default value: 0*
-   **v4l2-audio-input \** : Card input to use for audio *default value: 0*
-   **v4l2-io \** : IO method *default value: 0*
-   **v4l2-width \** : Prefered video width (if non zero) *default value: 0*
-   **v4l2-height \** : Prefered video height (if non zero) *default value: 0*
-   **v4l2-fps \** : Frames per second (if non zero) *default value: 0*

##### Audio input

These options do not apply to audio streams in compressed data.

-   **v4l2-adev \** : Audio input device *default value: ""*
-   **v4l2-audio-method \** : Allowed audio input methods (bitmask: 1 for OSS, 2 for ALSA) *default value: 3*
-   **v4l2-stereo** : Capture audio in stereo *default value: enabled*
-   **v4l2-samplerate \** : Audio input sample rate in Hz *default value: 48000*

##### Tuner

-   **v4l2-tuner \** : Tuner to use *default value: 0*
-   **v4l2-tuner-frequency \** : Tuner frequency in Hz or MHz depending on the underlying v4l2 driver *default value: -1*
-   **v4l2-tuner-audio-mode \** : Tuner audio mode *default value: -1*

##### Controls

These controls will be used only if they are supported by the v4l2 driver.

-   **v4l2-controls-reset** : Reset all the v4l2 controls *default value: disabled*
-   **v4l2-brightness \** : Brightness *default value: -1*
-   **v4l2-contrast \** : Contrast *default value: -1*
-   **v4l2-saturation \** : Saturation *default value: -1*
-   **v4l2-hue \** : Hue *default value: -1*
-   **v4l2-black-level \** : Black level *default value: -1*
-   **v4l2-auto-white-balance \** : Auto white balance *default value: -1*
-   **v4l2-do-white-balance \** : Do white balance *default value: -1*
-   **v4l2-red-balance \** : Red balance *default value: -1*
-   **v4l2-blue-balance \** : Blue balance *default value: -1*
-   **v4l2-gamma \** : Gamma *default value: -1*
-   **v4l2-exposure \** : Exposure *default value: -1*
-   **v4l2-autogain \** : Auto gain *default value: -1*
-   **v4l2-gain \** : Gain *default value: -1*
-   **v4l2-hflip \** : Flip the image horizontaly *default value: -1*
-   **v4l2-vflip \** : Flip the image verticaly *default value: -1*
-   **v4l2-hcenter \** : Horizontal center *default value: -1*
-   **v4l2-vcenter \** : Vertical center *default value: -1*
-   **v4l2-audio-volume \** : Audio volume *default value: -1*
-   **v4l2-audio-balance \** : Audio balance *default value: -1*
-   **v4l2-audio-mute** : Audio mute *default value: disabled*
-   **v4l2-audio-bass \** : Audio bass *default value: -1*
-   **v4l2-audio-treble \** : Audio treble *default value: -1*
-   **v4l2-audio-loudness \** : Audio loudness *default value: -1*
-   **v4l2-set-ctrls \** : Set any other control listed in the debug output using a comma seperated list in curly braces such as {video_bitrate=6000000,audio_crc=0,stream_type=3} *default value: ""*

#### Example

Open a video device with default settings:

    % vlc v4l2:///dev/video0:width=640:height=480

Get information about a video device's capabilities:

    % vlc -vvv --color v4l2:///dev/video0 --run-time 1 vlc://quit -I dummy -V dummy -A dummy

Command line for Hauppauge PVR 250 to get France 2 (at ECP) and encode as MPEG2 and stream using UDP multicast:

    % vlc -I dummy -vvv 'v4l2c://:audio-method=0:controls-reset:set-ctrls={video_bitrate_mode=1,video_bitrate=4000000,video_peak_bitrate=4000000}:width=720:height=576:tuner=0:tuner-frequency=478550'  --sout "#std{access=udp{ttl=12},mux=ts,url=239.255.1.1}"

Note: v4l2c is an alias used to force VLC to use the v4l2 module in it's Access variant without probing the Access Demux version first (the c stands for compressed).

#### Source code

-   [modules/access/v4l2/v4l2.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/v4l2/v4l2.c)

#### See also

-   Documentation:Modules/v4l
-   Documentation:Modules/dshow

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### vcd {#modules-vcd}

See also: VCD, SVCD and SVCD subtitles

Module: vcd

**Type**: Access

**First VLC version**: 0.5.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: VCD input

**Shortcut(s)**: 0, 1

#### Options

None.

#### Source code

-   [modules/access/vcd/vcd.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/vcd/vcd.c) (main file)
-   [modules/access/vcd](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/access/vcd) (folder)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### VHS {#modules-vhs}

Module: vhs

**Type**: Video filter

**First VLC version**: 2.2.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: VHS movie effect video filter

**Shortcut(s)**: -

#### Options

None.

#### Examples

Coloured stripes appear on-screen and the video shifts position after being paused.

    $ vlc --video-filter "vhs" video.ogv

#### Source code

-   [modules/video_filter/vhs.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_filter/vhs.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### visual {#modules-visual}

Module: visual

**Type**: Visualization

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Visualizer filter

**Shortcut(s)**: -

For the option 0 the values for 1{.variable} originate from [modules/visualization/window_presets.h](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/visualization/window_presets.h).

The code claims 0 must be within 1 ≤ 1{.variable} ≤ 127 but no checks seem to be performed.

The options 0, 1, 2 and 3 have been obsolete since VLC 1.0.0.

This module has a single shortcut: 0 (specifically with a 1{.sample}).

#### Options

-   **effects-list \ {dummy,scope,spectrum,spectrometer,vuMeter}** : A list of visual effect, separated by commas *default value: spectrum*
-   **effects-width \** : The width of the effects video window, in pixels *default value: 800*
-   **effects-height \** : The height of the effects video window, in pixels *default value: 500*
-   **effect-fft-window \ {none,hann,flattop,blackmanharris,kaiser}** : The type of FFT window to use for spectrum-based visualizations. Values correspond to "None", "Hann", "Flat Top", "Blackman-Harris", "Kaiser" *default value: flat*
-   **effect-kaiser-param \** : The parameter alpha for the Kaiser window. Increasing alpha increases the main-lobe width and decreases the side-lobe amplitude *default value: 3.0f*

##### Spectrum analyser

-   **visual-80-bands \** : Show 80 bands instead of 20 *default value: enabled*
-   **visual-peaks \** : Draw peaks in the analyzer *default value: enabled*

##### Spectrometer

-   **spect-show-original \** : Enable the "flat" spectrum analyzer in the spectrometer *default value: disabled*
-   **spect-show-base \** : Draw the base of the bands *default value: enabled*
-   **spect-radius \** : Defines radius size in pixels, of base of bands (beginning) *default value: 42*
-   **spect-sections \<integer \[1 .. 0{.variable}\]\>** : Determines how many sections of spectrum will exist *default value: 3*
-   **spect-color \** : YUV-Color cube shifting across the V-plane ( 0 - 127 ) *default value: 80*
-   **spect-show-bands \** : Draw bands in the spectrometer *default value: enabled*
-   **spect-80-bands \** : Show 80 bands instead of 20 *default value: enabled*
-   **spect-separ \** : Number of blank pixels between bands *default value: 1*
-   **spect-amp \** : This is a coefficient that modifies the height of the bands *default value: 8*
-   **spect-show-peaks \** : Draw peaks in the analyzer *default value: enabled*
-   **spect-peak-width \** : ***Additions or subtractions of pixels*** on the peak width *default value: 61*
-   **spect-peak-height \** : ***Total pixel height*** of the peak items *default value: 1*

#### Screenshots

Click thumbnails for larger images and author attribution.

-

    Spectrum

-

    Spectrometer

#### Source code

-   [modules/visualization/visual/visual.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/visualization/visual/visual.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### vobsub {#modules-vobsub}

Module: vobsub

**Type**: Access demux

**First VLC version**: 0.8.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Vobsub subtitles parser

**Shortcut(s)**: -

#### Options

None.

#### Source code

-   [modules/demux/vobsub.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/vobsub.c)
-   [modules/demux/vobsub.h](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/vobsub.h)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### vorbis {#modules-vorbis}

Module: vorbis

**Type**: Access demux

**First VLC version**: 0.5.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Vorbis audio decoder

**Shortcut(s)**: -

-   **sout-vorbis-quality \** : Enforce a quality between 1 (low) and 10 (high), instead of specifying a particular bitrate. This will produce a VBR stream
-   **sout-vorbis-max-bitrate \** : Maximum bitrate in kbps. This is useful for streaming applications
-   **sout-vorbis-min-bitrate \** : Minimum bitrate in kbps. This is useful for encoding for a fixed-size channel
-   **sout-vorbis-cbr \** : Force a constant bitrate encoding (CBR) *default value: disabled*

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### vpx {#modules-vpx}

This module has no shortcut.

Support for [DASH](http://en.wikipedia.org/wiki/Dynamic_Adaptive_Streaming_over_HTTP) in WebM is planned for VLC 4.0.0.

#### Demux

Module: vpx

**Type**: Access demux

**First VLC version**: 1.1.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: WebM video decoder

**Shortcut(s)**: -

VP9 support was added in VLC 2.1.1.

##### Options

None.

#### Mux

Module: vpx

**Type**: Muxer

**First VLC version**: 3.0.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: WebM video encoder

**Shortcut(s)**: -

##### Options

-   **sout-vpx-quality-mode \** : Quality setting which will determine max encoding time: 0 is Good quality, 1 is Realtime and 2 is Best quality *default value: 0*

#### Source code

-   [modules/codec/vpx.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/vpx.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### vsxu {#modules-vsxu}

Module: vsxu

**Type**: Visualization

**First VLC version**: 2.1.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Visualization module wrapper for Vovoid [VSXu](http://en.wikipedia.org/wiki/VSXu)

**Shortcut(s)**: -

#### Options

-   **vsxu-width \** : The width of the video window, in pixels *default value: 1280*
-   **vsxu-height \** : The height of the video window, in pixels *default value: 800*

#### Source code

-   [modules/visualization/vsxu.cpp](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/visualization/vsxu.cpp)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### wall {#modules-wall}

Module: wall

**Type**: Video output splitter

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: all

**Description**: Splits the video output in several windows

**Shortcut(s)**: -

You can use this module to split a video output in several small windows. This is especially useful if you want to display parts of the same video on several computers to make a big video wall.

The option 0 is [planned to be removed from 4.0.0](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=b75bde9e40dd0f7726a381c0dd2af571144b68ab) to fix [Bug #17433](https://trac.videolan.org/vlc/ticket/17433) and [Bug #213](https://trac.videolan.org/vlc/ticket/213). The option is redundant, and there will still be a way to select a custom ratio.

For the option 0, list the integers of the windows. To select windows 2, 3 and 5 specify --wall-active=2,3,5.

#### Options

-   **wall-cols \** : Number of horizontal windows in which to split the video *default value: 3*
-   **wall-rows \** : Number of vertical windows in which to split the video *default value: 3*
-   **wall-active \** : Comma-separated list of active windows, defaults to all *default value: NULL*
-   **wall-element-aspect \** : Aspect ratio of the individual displays building the wall *default value: 4:3*

#### Source code

-   [modules/video_splitter/wall.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_splitter/wall.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### wav {#modules-wav}

#### Demux

Module: wav

**Type**: Access demux

**First VLC version**: 0.5.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: WAV demuxer

**Shortcut(s)**: (none)

##### Options

None.

#### Mux

Module: wav

**Type**: Muxer

**First VLC version**: 0.8.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: WAV muxer

**Shortcut(s)**: 0

##### Options

None.

#### Source code

-   [modules/demux/wav.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/demux/wav.c)
-   [modules/mux/wav.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/mux/wav.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### wave {#modules-wave}

Module: wave

**Type**: Video filter

**First VLC version**: 0.9.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: wave video filter

**Shortcut(s)**: -

#### Example

**VLC 0.9.0 and above**:

    % vlc --video-filter wave somevideo.avi

**Note:** In versions prior to 0.9.0, wave was part of the distort video output filter.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### waveout {#modules-waveout}

Module: waveout

**Type**: Audio output

**First VLC version**: 0.4

**Last VLC version**: -

**Operating system(s)**: Windows

**Description**: Wave audio output

**Shortcut(s)**: -

#### Introduction

This is an optionnal audio output for Windows. It can be used when the DirectSound module doesn't work.

It uses the normal Multimedia Windows API that is present since Windows 95.

#### Options

##### Float32 output

This option allows you to enable or disable the high-quality float32 audio output mode (which is not well supported by some soundcards).

##### Audio Device Selection

This option allows you to select the audio output device listed.

It uses the string name of the audio device.

### wiimote {#modules-wiimote}

Module: wiimote

**Type**: Control

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: Enables to use the wiimote as a remote control

**Shortcut(s)**: -

Wiimote is used to control VLC Media player. Uses [CWiid](http://abstrakraft.org/cwiid/).

Example usage:

    % vlc --control wiimote

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### wingdi {#modules-wingdi}

Module: wingdi

**Type**: Video output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: Windows

**Description**: Windows GDI video output

**Shortcut(s)**: (none)

#### Source code

-   [modules/video_output/win32/wingdi.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/video_output/win32/wingdi.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### wxWidgets {#modules-wxwidgets}

**This page is obsolete and kept only for historical interest.** It may document features that are obsolete, superseded, or irrelevant. Do not rely on the information here being up-to-date.

WxWidgets was the default, plain, graphical, interface to VLC, made using the [WxWidgets](http://www.wxwidgets.org) library (Linux users may need to have this installed). It was used as the default interface on the Windows and Linux versions of VLC and have been replaced by the Qt Interface since 0.9.0.

It is known as the "wx" interface, so you can (or was able to) force this by running

    vlc -I wx

If WxWidgets is not available, it will probably revert to using the rc (console) interface, even if you force it. The most likely reason for this is if WxWidgets hasn't been installed, or if it wasn't linked in (using the ./configure). See compiling VLC [\[1\]](http://developers.videolan.org/vlc/) for information on compiling.

### x11 {#modules-x11}

Module: x11

**Type**: Video output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: X11 video output

**Shortcut(s)**: -

This shows video through [X11](http://en.wikipedia.org/wiki/X11). If you are using the wxwidgets or skins2 interface, the video will be shown inside the media player.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### x264 {#modules-x264}

This page is about the H.264 encoder. Were you looking for h26x, the H.264 decoder?

See also: Documentation:Modules/x265

Module: x264

**Type**: Muxer

**First VLC version**: 0.7.2

**Last VLC version**: -

**Operating system(s)**: all

**Description**: H.264/MPEG-4 Part 10/AVC encoder (x264)

**Shortcut(s)**: -

#### Options

##### Frame-type

-   **sout-x264-keyint \** : Maximum GOP size *default value: 250*
-   **sout-x264-min-keyint \** : Minimum GOP size *default value: 25*
-   **sout-x264-opengop \** : Use recovery points to close GOPs *default value: disabled*
-   **sout-x264-bluray-compat \** : Enable compatibility hacks for Blu-ray support *default value: disabled*
-   **sout-x264-scenecut \** : Extra I-frames aggressivity *default value: 40*
-   **sout-x264-bframes \** : B-frames between I and P *default value: 3*
-   **sout-x264-b-adapt \** : Adaptive B-frame decision *default value: 1*
-   **sout-x264-b-bias \** : Influence (bias) B-frames usage *default value: 0*
-   **sout-x264-bpyramid \ {none,strict,normal}** : Keep some B-frames as references *default value: "normal"*
-   **sout-x264-cabac \** : [CABAC](http://en.wikipedia.org/wiki/Context-adaptive_binary_arithmetic_coding) *default value: enabled*
-   **sout-x264-fullrange \** : Use fullrange instead of TV colorrange *default value: disabled*
-   **sout-x264-ref \** : Number of reference frames *default value: 3*
-   **sout-x264-nf \** : Skip loop filter *default value: disabled*
-   **sout-x264-deblock \** : Loop filter AlphaC0 and Beta parameters alpha:beta *default value: "0:0"*
-   **sout-x264-psy-rd \** : Strength of psychovisual optimization *default value: "1.0:0.0"*
-   **sout-x264-psy \** : Use Psy-optimizations *default value: enabled*
-   **sout-x264-level \** : H.264 level *default value: "0"*
-   **sout-x264-profile \** : H.264 profile *default value: "high"*
-   **sout-x264-interlaced \** : Interlaced mode *default value: disabled*
-   **sout-x264-frame-packing \ {-1,0,1,2,3,4,5,6}** : Frame packing *default value: -1*
-   **sout-x264-slices \** : Force number of slices per frame *default value: 0*
-   **sout-x264-slice-max-size \** : Limit the size of each slice in bytes *default value: 0*
-   **sout-x264-slice-max-mbs \** : Limit the size of each slice in macroblocks *default value: 0*
-   **sout-x264-hrd \** : HRD-timing information *default value: "none"*

##### Rate control

-   **sout-x264-qp \** : Set QP *default value: -1*
-   **sout-x264-crf \** : Quality-based VBR *default value: 23*
-   **sout-x264-qpmin \** : Min QP *default value: 10*
-   **sout-x264-qpmax \** : Max QP *default value: 51*
-   **sout-x264-qpstep \** : Max QP step between frames *default value: 4*
-   **sout-x264-ratetol \** : Average bitrate tolerance *default value: 1.0*
-   **sout-x264-vbv-maxrate \** : Max local bitrate *default value: 0*
-   **sout-x264-vbv-bufsize \** : VBV buffer *default value: 0*
-   **sout-x264-vbv-init \** : Initial VBV buffer occupancy *default value: 0.9*
-   **sout-x264-ipratio \** : QP factor between I and P *default value: 1.40*
-   **sout-x264-pbratio \** : QP factor between P and B *default value: 1.30*
-   **sout-x264-chroma-qb-offset \** : QP difference between chroma and luma *default value: 0*
-   **sout-x264-pass \ {0,1,2,3}** : Multipass ratecontrol *default value: 0*
-   **sout-x264-qcomp \** : QP curve compression *default value: 0.60*
-   **sout-x264-cplxblur \** : Reduce fluctuations in QP *default value: 20.0*
-   **sout-x264-qblur \** : Reduce fluctuations in QP *default value: 0.5*
-   **sout-x264-aq-mode \ {0,1,2}** : Defines bitdistribution mode for AQ, default 1 *default value: 0{.variable}*
-   **sout-x264-aq-strength \** : Strength of AQ *default value: 1.0*

##### Analysis

-   **sout-x264-partitions \ {none,fast,normal,slow,all}** : Partitions to consider *default value: "normal"*
-   **sout-x264-direct \ {none,spatial,temporal,auto}** : Direct MV prediction mode *default value: "spatial"*
-   **sout-x264-direct-8x8 \** : Direct prediction size *default value: 1*
-   **sout-x264-weightb \** : Weighted prediction for B-frames *default value: enabled*
-   **sout-x264-weightp \** : Weighted prediction for P-frames *default value: 2*
-   **sout-x264-me \ {dia,hex,umh,esa,tesa}** : Integer pixel motion estimation method *default value: "hex"*
-   **sout-x264-merange \** : Maximum motion vector search range *default value: 16*
-   **sout-x264-mvrange \** : Maximum motion vector length *default value: -1*
-   **sout-x264-mvrange-thread \** : Minimum buffer space between threads *default value: -1*
-   **sout-x264-subme \** : Subpixel motion estimation and partition decision quality *default value: 7*
-   **sout-x264-mixed-refs \** : Decide references on a per partition basis *default value: enabled*
-   **sout-x264-chroma-me \** : Chroma in motion estimation *default value: enabled*
-   **sout-x264-8x8dct \** : Adaptive spatial transform size *default value: enabled*
-   **sout-x264-trellis \ {0,1,2}** : Trellis RD quantization: This requires CABAC *default value: 1*
-   **sout-x264-lookahead \** : Framecount to use on frametype lookahead *default value: 40*
-   **sout-x264-intra-refresh \** : Use Periodic Intra Refresh *default value: disabled*
-   **sout-x264-mbtree \** : Use mb-tree ratecontrol *default value: enabled*
-   **sout-x264-fast-pskip \** : Early SKIP detection on P-frames *default value: enabled*
-   **sout-x264-dct-decimate \** : Coefficient thresholding on P-frames *default value: enabled*
-   **sout-x264-nr \** : Noise reduction *default value: 0*
-   **sout-x264-deadzone-inter \** : Inter luma quantization deadzone *default value: 21*
-   **sout-x264-deadzone-intra \** : Intra luma quantization deadzone *default value: 11*

##### Input/Output

-   **sout-x264-non-deterministic \** : Non-deterministic optimizations when threaded *default value: disabled*
-   **sout-x264-asm \** : CPU optimizations *default value: enabled*
-   **sout-x264-psnr \** : PSNR computation *default value: disabled*
-   **sout-x264-ssim \** : SSIM computation *default value: disabled*
-   **sout-x264-quiet \** : Quiet mode *default value: disabled*
-   **sout-x264-sps-id \** : SPS and PPS id numbers *default value: 0*
-   **sout-x264-aud \** : Access unit delimiters *default value: disabled*
-   **sout-x264-verbose \** : Statistics *default value: disabled*
-   **sout-x264-stats \** : Filename for 2 pass stats file *default value: "x264_2pass.log"*
-   **sout-x264-preset \** : Default preset setting used *default value: NULL*
-   **sout-x264-tune \** : Default tune setting used *default value: NULL*
-   **sout-x264-options \** : x264 advanced options, in the form 0 *default value: NULL*

##### Appendix

**For the option 0:**

-1
:   disabled

0
:   checkerboard - pixels are alternatively from L and R

1
:   column alternation - L and R are interlaced by column

2
:   row alternation - L and R are interlaced by row

3
:   side by side - L is on the left, R on the right

4
:   top bottom - L is on top, R on bottom

5
:   frame alternation - one view per frame

**For the option 0:**

1
:   First pass, creates stats file

2
:   Last pass, does not overwrite stats file

3
:   0{.variable}^th^ pass, overwrites stats file

**For the option 0:**

0
:   Disabled

1
:   Current x264 default mode

2
:   uses 0 instead of 1 and attempts to adapt strength per frame

**For the option 0:**

0
:   disabled

1
:   enabled only on the final encode of a MB

2
:   enabled on all mode decisions

#### Source code

-   [modules/codec/x264.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/x264.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### x265 {#modules-x265}

This page is about the H.265 encoder. Were you looking for h26x, the H.265 decoder?

See also: Documentation:Modules/x264

Module: x265

**Type**: Muxer

**First VLC version**: 2.2.0

**Last VLC version**: -

**Operating system(s)**: all

**Description**: H.265/HEVC encoder (x265)

**Shortcut(s)**: (none)

#### Options

None.

#### Source code

-   [modules/codec/x265.c](https://git.videolan.org/?p=vlc.git;a=blob;f=modules/codec/x265.c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### xvideo {#modules-xvideo}

Module: xvideo

**Type**: Video output

**First VLC version**: -

**Last VLC version**: -

**Operating system(s)**: Linux

**Description**: XVideo video output

**Shortcut(s)**: -

This is the default output for Linux systems, showing video through the [XVideo](http://en.wikipedia.org/wiki/XVideo) X Window System extension. If you are using the wxwidgets or skins2 interface, the video will be shown inside the media player.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

## Playing Media {#playing-media}

### Advanced Use of VLC {#play-howto-advanced-use-of-vlc}

#### Use the command line

**TODO: completely outdated**

This page is outdated and information might be incorrect.

All standard operations of VLC should be available from the GUI. However, some complex operations can only be done from the command line and there are situations in which you don't need or want a GUI. Here is the complete description of VLC's command line and how to use it.

You need to be quite comfortable with command line usage to use this.

    Note: Windows users have to use the --option-name="value" syntax instead of the --option-name value syntax.

##### Getting help

VLC uses a modular structure. The core mainly manages communication between modules. All the multimedia processing is done by modules. There are input modules, demultiplexers, decoders, video output modules, ...

This chapter will only describe the "general" options, i.e. the core options. Each module adds new options. For example, the HTTP input module will add options for caching, proxy, authentication, ...

By using **vlc --help**, you will get the basic core options. **vlc --longhelp** will give all the basic options (core + modules). Adding **--advanced** will give the "advanced options" (for advanced users). So **vlc --longhelp --advanced** will give you all options. You can also append **--help-verbose** if you want more detailed help.

Also, you might want to get debug information. To do this, use **-v** or **-vv** (this will show lower severity messages). If your console supports it, you can add **--color to get messages in color.**

##### Opening streams

The following commands start VLC and start reading the given element(s):

###### Opening a file

Start VLC with:

    % vlc my_file

VLC should be able to recognize the file type. If it does not, you can force demultiplexer and decoder (see below).

A list of all video and audio codecs supported by VLC check the [VLC features list](https://www.videolan.org/vlc/features.html).

###### Opening a DVD or VCD, or an audio CD

Start VLC with:

For a DVD with menus:

    % vlc dvd://[device][@raw_device][@[title][:[chapter][:angle]]]

In most cases, **vlc dvd://** or **vlc dvd://\[device\]** will do. \[device\] is for example **/dev/dvd** on GNU/Linux or **D:** on Windows (complete path to your DVD drive).

or

(DVD without menus):

    % vlc dvdsimple://[device][@raw_device][@[title][:[chapter][:angle]]]

or

(VCD):

    % vlc vcd://[device][@{E|P|E|T|S}[number]]

or

(Audio CD):

    % vlc cdda://[device][@[track]]

###### Receiving a network stream

To receive an unicast RTP/UDP stream (sent by VLC's stream output), start VLC with:

    % vlc rtp://@:5004

If 5004 is the port to which packets are sent. 1234 is another commonly used port number. you use the default port (1234), **vlc rtp://** will do. For more information, look at the Streaming Howto.

To receive an multicast UDP/RTP stream (sent by VLC's stream output), start VLC with:

    % vlc rtp://@multicast_address:port

To receive a SSM (source specific multicast) stream, you can use:

    % vlc rtp://server_address@multicast_address:port

This only works on Operating systems that support SSM (Windows XP and Linux).

To receive a HTTP stream, start VLC with:

    % vlc 0

To receive a RTSP stream, start VLC with:

    % vlc rtsp://www.example.org/your_stream

##### Modules selection

VLC always tries to select the most appropriate interface, input and output modules, among the ones available on the system, according to the stream it is given to read. However, you may wish to force the use of a specific module with the following options.

-   **--intf \** allows you to select the interface module.
-   **--extraintf \** allows you to select extra interface modules that will be launched in addition to the main one. This is mainly useful for special *control* interfaces, like HTTP, RC (Remote Control), ... (see below)
-   **--aout \** allows you to select the audio output module.
-   **--vout \** allows you to select the video output module.
-   **--memcpy \** allows you to choose a memory copy module. You should probably never touch that.

You can get a listing of the available modules by using **vlc -l**

##### Stream Output

The Stream output system allows vlc to become a streaming server.

For more details on the stream output system, please have a look at the [Streaming HowTo](#streaming-howto).

##### Other Options

###### Audio options

-   **--noaudio** disables audio output. Note that if you are streaming (ex: to a file) this has no effect (streaming copies the audio verbatim). Use --sout-xxx instead (ex: --no-sout-audio)
-   **--mono** forces VLC to treat the stream in mono audio.
-   **--volume \** sets the level of audio output (between 0 and 1024). Also only applies to local playback (like --noaudio).
-   **--aout-rate \** sets the audio output frequency (Hz). By default, VLC will try to autodetect this.
-   **--desync \** compensates desynchronization of audio (ms). (If audio and video streams are not synchronized, use this setting to delay the audio stream)
-   **--audio-filter \** adds audio filters to the processing chain. Available filters are visual (visualizer with spectrum analyzer and oscilloscope), headphone (virtual headphone patialization) and normalizer (volume normalizer)

###### Video options

-   **--no-video** disables video output.
-   **--grayscale** turns video output into grayscale mode.
-   **--fullscreen** ( or **-f**) sets fullscreen video.
-   **--nooverlay** disables hardware acceleration for the video output.
-   **--width, --height \** sets the video window dimensions. By default, the video window size will be adjusted to match the video dimensions.
-   **--start-time \** starts the video here; the integer is the number of seconds from the beginning (e.g. 1:30 is written as 90)
-   **--stop-time \** stops the video here; the integer is the number of seconds from the beginning (e.g. 1:30 is written as 90)
-   **--zoom \** adds a zoom factor.
-   **--aspect-ratio \** forces source aspect ratio. Modes are 4x3, 16x9, ...
-   **--spumargin \** forces SPU subtitles position.
-   **--video-filter \** adds video filters to the processing chain. You can add several filters, separated by commas
-   **--sub-filter \** adds video subpictures filter to the processing chain.

###### Desktop/Screen grab options

You can see the various options for "grabbing the desktop" (VLC's built-in screen grabber capture device) by using the GUI. See 0

###### Playlist options

-   **--random** plays files randomly forever.
-   **--loop** loops playlist on end.
-   **--repeat** repeats current item until another item is forced
-   **--play-and-stop** stops the playlist after each played item.

###### Network options

-   **--server-port \** sets server port.
-   **--iface \** specifies the network interface to use.
-   **--iface-addr \** specifies your network interface IP address.
-   **--mtu \** specifies the MTU of the network interface.
-   **--ipv6** forces IPv6.
-   **--ipv4** forces IPv4.

###### CPU options

You should probably not touch these options unless you know what you are doing.

-   **--nommx** disables the use of MMX CPU extensions.
-   **--no3dn** disables the use of 3D Now! CPU extensions.
-   **--nommxext** disables the use of MMX Ext CPU extensions.
-   **--nosse** disables the use of SSE CPU extensions.
-   **--noaltivec** disables the use of Altivec CPU extensions.

###### Miscellaneous options

-   **--quiet** deactivates all console messages.
-   **--color** displays color messages.
-   **--search-path \** specifies interface default search path.
-   **--plugin-path \** specifies plugin search path.
-   **--no-plugins-cache** disables the plugin cache (plugins cache speeds up startup)
-   **--dvd \** specifies the default DVD device.
-   **--vcd \** specifies the default VCD device.
-   **--program \** specifies program (SID) (for streams with several programs, like satellite ones).
-   **--audio-type \** specifies the default audio type to use with dvds.
-   **--audio-channel \** specifies the default audio channel to use with dvds.
-   **--spu-channel \** specifies the default subtitle channel to use with dvds.
-   **--version** gives you information about the current VLC version.
-   **--module \** displays help about specified module. (Shortcut: **-p**)

##### Item-specific options

There are many options that are related to items (like **--novideo**, **--codec**, **--fullscreen**).

For all of these, you have the possibility to make them item-specific, using ":" instead of "--" and putting the option just after the concerned item.

Examples:

    % vlc file1.mpg :fullscreen file2.mpg

will play file1.mpg in fullscreen mode and file2.mpg in the default mode (which is generally no fullscreen), whereas

    % vlc --fullscreen file1.mpg file2.mpg

will play both files in fullscreen mode

    % vlc --fullscreen file1.mpg :sub-file=file1.srt :no-fullscreen file2.mpg :filter=distort

will play file1.mpg in windowed (no-fullscreen) mode with the subtitles file file1.srt and will play file2.mpg with video filter distort enabled in fullscreen mode (item-specific options override global options).

#### Advanced use of filters

##### Filters

These are the old style VLC filters. They only apply to on screen display and thus cannot be streamed. However, on version 1.1.11 you are still able to apply these filters in *transcode* module using parameter *vfilter*. More information can be found on Documentation:Streaming HowTo/Advanced Streaming Using the Command Line#vfilter.

###### Deinterlacing video filter

Module name: **deinterlace**

-   **--deinterlace-mode {discard,blend,mean,bob,linear,x,yadif,yadif (2x),phosophor,ivtc}** choose a deinterlacing mode.

###### Invert video filter

Module name: **invert**

###### Image properties filter

Module name: **adjust**

*[Transcluded](http://en.wikipedia.org/wiki/mw:Transclusion) from Documentation:Modules/adjust*

-   **contrast \** : Contrast *default value: 1.0*
-   **brightness \** : Brightness *default value: 1.0*
-   **hue \** : Hue *default value: 0*
-   **saturation \** : Saturation *default value: 1.0*
-   **gamma \** : Gamma *default value: 1.0*
-   **brightness-threshold \** : When this mode is enabled, pixels will be shown as black or white. Also may invert the brightness value. The threshold value will be the brightness defined below *default value: disabled*

###### Wall video filter

Module name: **wall** This filter splits the output in several windows.

*[Transcluded](http://en.wikipedia.org/wiki/mw:Transclusion) from Documentation:Modules/wall*

-   **wall-cols \** : Number of horizontal windows in which to split the video *default value: 3*
-   **wall-rows \** : Number of vertical windows in which to split the video *default value: 3*
-   **wall-active \** : Comma-separated list of active windows, defaults to all *default value: NULL*
-   **wall-element-aspect \** : Aspect ratio of the individual displays building the wall *default value: 4:3*

###### Video transformation filter

Module name: **transform**

*[Transcluded](http://en.wikipedia.org/wiki/mw:Transclusion) from Documentation:Modules/transform*

-   **transform-type \ { "90", "180", "270", "hflip", "vflip", "transpose", "antitranspose" }** : Transformation type *default value: "90"*

###### Distort video filter

Module name: **distort**

*See Documentation:Modules/distort*

###### Clone video filter

This filter clones the output window.

Module name: **clone**

*[Transcluded](http://en.wikipedia.org/wiki/mw:Transclusion) from Documentation:Modules/clone*

-   **clone-count \** : Number of video windows in which to clone the video. *default value: 2*
-   **clone-vout-list \** : You can use specific video output modules for the clones. Use a comma-separated list of modules. *default value: ""*

###### Croppadd video filter

Module name: **croppadd**

*[Transcluded](http://en.wikipedia.org/wiki/mw:Transclusion) from Documentation:Modules/croppadd*

-   **croppadd-croptop \<integer \[0 .. 0{.variable}\]\>** : Pixels to crop from top
-   **croppadd-cropbottom \<integer \[0 .. 0{.variable}\]\>** : Pixels to crop from bottom
-   **croppadd-cropleft \<integer \[0 .. 0{.variable}\]\>** : Pixels to crop from left
-   **croppadd-cropright \<integer \[0 .. 0{.variable}\]\>** : Pixels to crop from right
-   **croppadd-paddtop \<integer \[0 .. 0{.variable}\]\>** : Pixels to add to top
-   **croppadd-paddbottom \<integer \[0 .. 0{.variable}\]\>** : Pixels to add to bottom
-   **croppadd-paddleft \<integer \[0 .. 0{.variable}\]\>** : Pixels to add to left
-   **croppadd-paddright \<integer \[0 .. 0{.variable}\]\>** : Pixels to add to right

###### Motion blur filter

Module name: **motionblur**

*[Transcluded](http://en.wikipedia.org/wiki/mw:Transclusion) from Documentation:Modules/motionblur*

-   **motionblur-factor \** : The bluring factor (1 to 127). Higher values mean more blurring *default value: 80*

###### Video pictures blending

Module name: **blend**

###### Video scaling filter

Module name: **scale**

##### Subpictures Filters

These are the new VLC filters. They can be streamed.

###### Marquee display sub filter

Module name: **marq**

*[Transcluded](http://en.wikipedia.org/wiki/mw:Transclusion) from Documentation:Modules/marq*

-   **marq-marquee \** : Marquee text to display. *default value: VLC*
-   **marq-file \** : File to read the marquee text from. *default value: NULL*
-   **marq-x \** : X offset, from the left screen edge. *default value: 0*
-   **marq-y \** : Y offset, down from the top. *default value: 0*
-   **marq-position \** : Marquee position: 0=center, 1=left, 2=right, 4=top, 8=bottom, you can also use combinations of these values, eg 6 = top-right. *default value: -1*
-   **marq-opacity \** : Opacity (inverse of transparency) of overlaid text. 0 = transparent, 255 = totally opaque. *default value: 255*
-   **marq-color \ { 0x000000, 0x808080, 0xC0C0C0, 0xFFFFFF, 0x800000, 0xFF0000, 0xFF00FF, 0xFFFF00, 0x808000, 0x008000, 0x008080, 0x00FF00, 0x800080, 0x000080, 0x0000FF, 0x00FFFF }** : Color of the text that will be rendered on the video. This must be an hexadecimal (like HTML colors). The first two chars are for red, then green, then blue. *default value: 0xFFFFFF*
-   **marq-size \** : Font size, in pixels. 0 uses the default font size. *default value: 0*
-   **marq-timeout \** : Number of milliseconds the marquee must remain displayed. 0 means forever. *default value: 0*
-   **marq-refresh \** : Number of milliseconds between string updates. This is mainly useful when using meta data or time format string sequences. *default value: 1000*

###### Logo video filter

Module name: **logo**

*[Transcluded](http://en.wikipedia.org/wiki/mw:Transclusion) from Documentation:Modules/logo*

-   **logo-file \** : Image to display. The full format is 0.
-   **logo-x \** : X offset from upper left corner. *default value: 0*
-   **logo-y \** : Y offset from upper left corner. *default value: 0*
-   **logo-position \ { 0, 1, 2, 4, 8, 5, 6, 9, 10 }** : Logo position. *default value: 5*
-   **logo-opacity \** : Logo opacity. 0 is transparent, 255 is fully opaque. *default value: 255*
-   **logo-delay \** : Global delay in [ms](http://en.wiktionary.org/wiki/ms#Translingual). Sets the duration each image will be displayed for in a loop iteration unless specified otherwise in the 0 option. *default value: 1000*
-   **logo-repeat \** : Number of loops for the logo animation. -1 for continuous, 0 to disable. *default value: -1*

This filter can be used both as an old style filter or a subpictures filter.

Note: You can move the logo by left-clicking on it.

#### The HTTP interface

VLC ships with a little HTTP server integrated. It is used both to stream using HTTP, and for the HTTP remote control interface.

To start VLC with the HTTP interface, use:

    % vlc -I http [--http-src /directory/] [--http-host host:port]

If you want to have both the "normal" interface and the HTTP interface, use **vlc --extraintf http**.

The HTTP interface will start listening at host:port (\:8080 if omitted), and will reproduce the structure of /directory at 0 ( vlc_source_path/share/http if omitted ).

Use a browser to go to 0. You should be taken to the main page.

VLC is shipped with a set of files that should be enough for generic needs. It is also possible to customize pages. See Documentation:Play HowTo/Building Pages for the HTTP Interface.

Available pages for 1.0.3 :

-   http://host:port - Main Interface
-   0 - VLM Interface
-   0 - Mosaic Wizard
-   0 - Flash based remote playback

#### Other control interfaces

VLC includes a number of so-called interfaces that are not really interfaces, but means of. Nevertheless, they are enabled by setting them as interface or extra interface, either in the Preferences, in General/Interface, or using **-I** or **--extraintf** on the command line.

##### Hotkeys

This module allows you to control VLC and playback via hotkeys. It is always enabled by default. You can use hotkeys in the video output window, you can't in the audio dummy interface.

Hotkeys can be hacked by:

    % vlc --key-

Code is composed by modifiers keys (Alt, Shift, Ctrl, Meta,Command) separated by a dash (-) and terminated by a key (a...z, +, =, -, ',', +, \<, \>, \`, /, ;, ', \\, \[, \], \*, Left, Right, Up, Down, Space, Enter, F1...F12, Home, End, Menu, Esc, Page Up, Page Down, Tab, Backspace, Mouse Wheel Up and Mouse Wheel Down). Main controls are available from hotkeys, such as : fullscreen, play-pause, faster, slower, next, prev, stop, quit, vol-up, etc. (use the **--longhelp** option for full list of functions). For example, for binding fullscreen to Ctrl-f, run:

    % vlc --key-fullscreen 'Ctrl-f'

The list of the default hotkeys is available here.

##### RC and RTCI

These two interfaces allow you to control VLC from a command shell (possibly using a remote connexion or a Unix socket).

Start VLC with **-I rc** or **--extraintf rc**. When you get the **Remote control interface initialized, \`h' for help** message, press h and Enter to get help about available commands.

To be able to remote connect to your VLC using a TCP socket (telnet-like connexion), use **--rc-host your_host:port**. Then, by connecting (using telnet or netcat) to the host on the given port, you will get the command shell.

To use a UNIX socket (local socket, this does not work for Windows), use **--rc-unix /path/to/socket**. Commands can then be passed using this UNIX socket.

The RTCI interface gives you more advanced options, such as marquee control for the marquee subpicture filter (See filter section).

##### Ncurses

This is a text interface, using ncurses library.

Start VLC with **-I ncurses** or **--extraintf ncurses**.

The ncurses interface

Press h to get the list of all available commands, with a short description.

There is also a filebrowser available for the ncurses interface in order to add playlist items. Press 'B' to use it.

The ncurses filebrowser

You can set the filebrowser starting point by launching vlc with the **--browse-dir** option:

    % vlc -I ncurses --browse-dir /filebrowser/starting/point/

##### Gestures

Gestures provide a simple mouse gestures control. TODO

#### The Mozilla plugin

VLC can also be embedded in a web browser! The following browsers are supported: [Firefox](https://www.mozilla.org/products/firefox/) and [Safari](https://www.apple.com/macosx/features/safari).

##### Install the plugin

###### GNU/Linux Debian, Ubuntu, etc.

Install the *mozilla-plugin-vlc* package using your preferred package manager. For example, at the command line enter:

    # apt-get update
    # apt-get install mozilla-plugin-vlc

###### Windows

Quit Firefox or Mozilla.

Select the Mozilla Plugin option when installing VLC Media Player. The installer will then automatically detect your browser and install the plugin.

Restart Firefox or Mozilla.

###### Manual Install

In ["Mozilla Firefox\\plugins"](http://kb.mozillazine.org/Installation_directory)

Create the directory if it doesn't exist.

**Folders** to copy:

-   osdmenu
-   plugins

**Files** to copy:

-   vlc.exe
-   vlc.exe.manifest
-   vlc-cache-gen.exe
-   npvlc.dll.manifest
-   npvlc.dll
-   libvlccore.dll
-   libvlc.dll
-   libvlc.dll.manifest
-   axvlc.dll
-   axvlc.dll.manifest

###### macOS

*The Mozilla/Safari plugin for macOS is only available from vlc version 0.8.5.1 and onwards.*

Quit Safari browser.

Download the Mozilla/safari plugin package from [macOS download page](https://www.videolan.org/vlc/download-macosx.html).

Run the installer from the dmg image.

###### Compile the sources yourself

Please look at the [developers page](https://www.videolan.org/developers) for information on how to do this.

##### Use the Mozilla plugin

If in the browser you open a link to an audio or video URL handled by the VLC plugin, or if a web page has HTML code that embeds audio or video handled by the VLC plugin, then the plugin should start and play the audio/video. Note the plugin (as of version 1.1.9) does not present any user interface — it has no default control panel and no keyboard shortcuts.

To get the list of the media types handled by the VLC plugin, browse to **about:plugins**. Conflicts will arise if you have more than one plugin installed that supports the same media type.

See the [Web plugin documentation](#webplugin) to create HTML pages that use JavaScript to control the plugin.

#### Snapshot Tool

Did you know you can use special codes to automatically generate filenames in the Snapshot Tool?

#### Specifying Streaming Options

*Further information: Documentation:Streaming HowTo New*

#### Audio Bar Graph over Video

This section specifies how to enable the audiobargraph audio filter and video overlay, (mostly) via the GUI. This displays an audio meter overlaid on the video.

There are three parts - an audio filter, which sends it's output via TCP to the Remote Control (RC) Interface. This information is then picked up and displayed by the Audio Bar Graph video subpicture filter (OSD).

To enable this, VLC needs to be started with the **--rc-host** command-line switch - e.g.

    % "C:\Program Files\VideoLAN\VLC\vlc.exe" --rc-host localhost:12345

In the GUI, set the following (this example from VLC v1.1.9 on Windows 7):

-   Preferences:Show settings:All
-   Audio/Filters \> Enable "Audio part of the BarGraph function"
-   Audio/Filters/audiobargraph \> use defaults, change "Sends the barGraph information every n audio packets" to 1 to enable see a more accurate display
-   Interface/Main interfaces \> Enable "Remote control interface"
-   Interface/Main interfaces/RC \> Enable "Do not open a DOS command box interface"
-   Video/Subtitles-OSD \> Enable "Audio Bar Graph Video sub filter"
-   Video/Subtitles-OSD/Audio Bar Graph \> Set the following settings:
    -   "Value of the audio channels levels" = 0 (setting this to 0:1 crashes VLC v1.1.9)
    -   "X coordinate" = 0
    -   "Y coordinate" = 0 (this doesn't seem to affect anything)
    -   "Transparency of the bargraph" = 128 for 50% transparency which looks ok
    -   "Bargraph position" = Left (seems to only work Left,Center,Right - can't go top or bottom)
    -   "Alarm" = 1 (enables the silence alarm - puts a red border around the bargraph if silent for too long)
    -   "Bar width in pixel" = 10 (20 if you want it to be really visible)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use {#play-howto-basic-use}

**Languages:English** • Deutsch

Read on for a quick overview of VLC's features and capabilities.

#### Starting VLC

##### Windows

-   In Windows XP: Click **Start** -\> **Programs** -\> **VideoLAN** -\> **VLC media player**.
-   In Windows 7: Click **Start** -\> **All Programs** -\> **VideoLAN** -\> **VLC media player**.

VLC is shown on the screen and a small icon  is shown in the system tray.

##### macOS

Start VLC from the applications menu or the system dock.

VLC is shown on the screen and a small icon  is shown in the dock.

##### Linux

Start VLC from the command line with **vlc** or start it from your desktop environment's application launcher.

#### Interface

##### The main interface

-   VLC media player on Windows and Linux: **VLC media player on macOS**

##### More interface informations

Go to [Documentation:Interface](#interface)

#### Play a media

##### Play a single media file

Find a media file you want to play with your favourite File Explorer (Windows Explorer, Finder, Konqueror...) and double-click on it.

You can also drag and drop the file onto VLC.

##### Play a whole media folder

Start VLC, open the *Media* menu, and select the *Open Folder...* menu item. An *Open Folder* dialog box will appear. Select the folder you want to open and select *Open*.

##### Play a CD/DVD/VCD

Insert your disk and your OS should ask you what you want to do. Select *Play with VLC* and select the OK button.

##### More open options

Go to [Documentation:Open Media](#open-media)

#### Preferences

##### Where are the VLC preferences?

To open the *Preferences* panel, open the Tools menu , and select the *Preferences* menu item.

Here is the Simple Preferences panel where you can modify the essential settings of VLC.

##### How to reset the VLC preferences?

Go to VSG:ResetPrefs

#### Playlist view

##### Overview

This view allows you to easily browse different sources of media. To access the Playlist View, click on the *Playlist* button in the main interface.

-   **1:**: The current Playlist you are listening and your Media Library

-   **2:**: The OS default media folders

-   **3:**: Your local optic drive (CD, DVD...)

-   **4:**: Your local network sources

-   **5:**: Internet sources (Podcasts, Shoutcast radios...)

-   **6:**: The media listing you are listening or browsing


##### More Playlist options

Go to [Documentation:Playlist](#playlist)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use 0.8 {#play-howto-basic-use-0-8}

**This page is obsolete and kept only for historical interest.** It may document features that are obsolete, superseded, or irrelevant. Do not rely on the information here being up-to-date.

**Note: this documentation is for versions older than 0.9. For help with 0.9 please see [VLC for dummies: an introduction to VLC](#vlc-for-dummies) and Basic Use for 0.9.**

#### General interface description

VLC has several interfaces:

-   A cross-platform interface, for Windows and GNU/Linux, called wxWidgets,
-   A native Mac OS X interface, and
-   A skinnable interface for Windows and GNU/Linux.

Screenshots below are drawn from the various interfaces, but VLC's functions work essentially the same on all operating systems.

##### Windows and GNU/Linux (wxWidgets)

This is the default interface on Windows and GNU/Linux (the screenshot is done on GNU/Linux, but it would look quite the same on Windows).

 The wxWidgets interface

This interface also features an *Extended GUI* which contains many additional features. To display or hide it, go to the *Settings* menu and click *Extended GUI*.

 The wxWidgets interface with extended GUI

##### Native Mac OS X (Cocoa)

This is the default interface on Mac OS X.

 The Mac OS X interface

This interface features an *Extended GUI* as well. It is called "Extended Controls" and can be opened through the Window menu.

 The Mac OSX interface with with the extended controls panel

#### Basic playback

##### Play a file

To play a file, open the *File* menu, and select the *Quick Open File* menu item. An Open File dialog box will appear. Select the file you want to open, and select Open. VLC will start playing the selected file.

An alternative is to drag 'n' drop your file on the VLC main interface or playlist window from the file explorer (Finder on MacOS X).

 The File menu - wxWidgets interface

 The File menu - MacOS X interface

 The Open file dialog - wxWidgets interface

 The Open file dialog - MacOS X interface

##### Play a CD/DVD/VCD

To Play a CD, VCD or a DVD, open the *File* menu, and select the *Open Disc...* menu item. In the Open Disk Dialog Box, select the type of media (DVD, VCD or Audio CD). When reading a DVD, you can enable DVD menus by selecting the *DVD (menus)* disc type in the wxWidgets Interface. In the MacOS X interface, this can be done by selecting the "Use DVD menus" dialog box.

You can select the drive from which the media should be read by giving the appropriate drive letter or device name in the "Device Name" text input. This should be auto-detected on MacOS X.

If you want to start the DVD or VCD playback from a given title and chapter instead of from the beginning, you can set it using the *Title* and *Chapter* selectors.

You can start playback by selecting the *Ok* button.

 The Open disk dialog - wxWidgets interface

 The Open disk dialog - MacOS X interface

##### Play a network stream (WebRadio, WebTV, etc.)

To open a network stream, open the "File" menu and select the "Open Network Stream" menu item.

-   To open a UDP unicast stream, select *UDP/RTP*, and set the appropriate UDP port in the selector (it is 1234 for streams sent by a VLC or VLS server).
-   To open a UDP multicast stream, select *UDP/RTP multicast*. Give the address of the multicast group in the "Address" text input, and select the appropriate UDP port.
-   To open a stream sent over http (Webradios, WebTVs, Shoutcast, Icecast...), ftp, or mms (Microsoft Media Server), select "HTTP/FTP/MMS", and give the corresponding complete URL, (such as 0 or [mms://live.ms.stream.net:8080/live.asf](mms://live.ms.stream.net:8080/live.asf)) in the corresponding text input. This also the way to open a RTSP stream with the MacOS X interface.
-   To open a RTSP stream (sent by Darwin Streaming Server, VLC, etc), in the wxWidgets interface, select "RTSP" and give the URL in the text input.

You can start playback by selecting the *Ok* button.

If you get some stuttering during playback, you can try to increase the size of the read buffer. This can be done in the *Open Network Stream* dialog box, by selecting the *Caching* box. You can then choose the amount time (in milliseconds) VLC should store data in its buffer before starting playback.

 The Open network dialog - wxWidgets interface

 The Open network dialog - MacOS X interface

##### Play from an acquisition card

This currently only possible on Linux and Windows. Open the File menu, and select "Open Capture Device..."

On Windows, supported cards include webcams, TV cards, acquisition cards... provided they come with directshow compatible drivers (Almost all acquisition cards do). You can choose the device to use for video and audio capture using the "Video device name" and "Audio device name" selectors. If your device doesn't appear in the list, try to select the "Refresh list" button. You can access the settings of your acquisition device by selecting the *configure* button. Options here depend on the driver of the device. You can select the "Device Proprieties" box if you want the configuration dialog box of every device to be displayed after having pressed the *Ok* button. Select the *Tuner properties* box to be prompted for tuner settings (PAL/NTSC standard, frequency...) for TV cards. The *Advanced options...* button allows to select some further settings useful in some rare cases, such as the chroma of the input (the way colors are encoded) and the size of the input buffer.

 The Open Capture device dialog and a device configuration windows- wxWidgets interface

On Linux, supported cards include webcams, TV cards, acquisition cards, provided they are supported by the Video4Linux architecture. Haupaugge PVR 250/350 cards are also supported, using the [IVTV drivers](http://ivtv.sourceforge.net/).

-   For Video4Linux devices, you can set the name of the video and audio devices using the "Video device name" and "Audio device name" text inputs. The "Advanced options..." button allows to select some further settings useful in some rare cases, such as the chroma of the input (the way colors are encoded) and the size of the input buffer.

 The Open Video4Linux dialog- wxWidgets interface

-   To use a Hauppauge PVR card, select the PVR tab in the "Open" dialog box. Use the "Device" text input to set the device of the card you want to use. You can set the Norm of the tuner (PAL, SECAM or NTSC) by using the "Norm" Drop Down. The Frequency selector allows you to set the frequency of the tuner (in kHz), the bitrate selector to set the bitrate of the resulting encoded stream (in bit/s). The "Advanced Options button allows to set some more settings, such as the size of the encoded video (in pixels), its framerate (in frame per second), the interval between 2 key frames, etc.

After having set all the required parameters, you can start the capture by selecting the "Ok" button.

 The Open PVR dialog- wxWidgets interface

#### Playlist

VLC can store a list of several files to play one after the other, using its playlist system. To access the playlist, click on the *Playlist* button on the main interface.

Each time you use the Open dialog box, the stream you select is appended at the end of the playlist and started.

The playlist window shows all the available streams. Double-click one to play it.

 The Playlist - wxWidgets interface

 The Playlist - MacOS X interface

##### Adding items, saving and loading playlists

In the wxWidgets interface, the *Manage* menu allows you to append an item at the end of the playlist (its playback won't start immediately), to save the playlist as a M3U or PLS file, or to import a playlist file.

In the MacOS X interface, saving a playlist can be done using the *Save Playlist...* function in the *File* menu. To import a playlist file, open it the same way as any other media file, using the *Quick Open File...* menu item.

##### Sorting

In the wxWidgets interface, *Sort* allows you to sort the playlist according to several criteria, or to shuffle it. You can also sort by clicking the header of the column.

In the MacOS X interface, sorting can be done by clicking the header of the column matching the criteria you want to use for sorting.

##### Playlist modes

The playlist supports several playback modes.

In the wxWidgets interface, the toolbar contains three playlist mode buttons. They allow to enable random mode, to repeat the whole playlist or to repeat one item.

In the MacOS X interface, random mode can be enabled by selecting the *Random* box. A drop down menu allows you to enable playlist and item repeat modes.

##### Misc

###### Search

You also have a search tool. Enter a search string and hit search. The next item to match the string will be highlighted. Keep hitting Search to cycle between all matching items.

###### Moving items

In the wxWidgets interface, the *Up* and *Down* buttons at the bottom of the playlist window allow you to move an item. Select an item and use these buttons to move it.

In the MacOS X interface, you can easily move an item with the mouse, using drag-and-drop.

###### Contextual menu

By right-clicking or control-clicking an item, a contextual menu will appear, giving access to a number of functions (for example, play the item, disable it, delete it, or get info on it).

If you ask for info, an *item info* dialog box will appear. This dialog box also allows you to change the name, the author and the location of the item to play.

 Item Info Dialog - wx Interface

 Item Info Dialog - MacOS X interface

#### Subtitles

VLC supports many kinds of subtitles.

##### Media with included subtitles

Many types of media can have embedded subtitles. VLC can read subtitles for the following media:

-   DVD
-   SVCD
-   OGM files
-   Matroska (MKV) files

Subtitles are disabled by default. To enable them, go to the *Video* menu, and to *Subtitles track*. All available subtitles tracks will be listed. Select one to get the subtitles. Depending on the media, a description (language, for example) might be available for the track.

 Select a subtitles track under Windows or Linux

 Select a subtitles track under MacOS X

DVD and SVCD subtitles are merely images, so you won't be able to change anything for them. OGM and Matroska subtitles are rendered text, so you will be able to change several options.

Text rendering options can be changed in the Preferences. In the *Modules* section, *text renderer* subsection, open the *freetype* page. You can then set the font and its size. For the font, you have to select a font file. Under Windows, they can be found in *C:\\Windows\\Fonts*. Under MacOS X, they are in */System/Library/Fonts*. Size can be set either relatively or as a number of pixels.

You need to restart your stream for the font modifications to take effect.

##### Subtitles files

While modern file formats like Matroska or OGM can handle subtitles directly, older formats like AVI can't. Therefore, a number of subtitles files formats have been created. You need two files: the video file and the subtitles files that only contains the text of the subtitles and timestamps.

VLC can handle these types of subtitles files:

-   MicroDVD
-   SubRIP
-   SubViewer
-   SSA
-   Sami
-   Vobsub (this one is quite special: it is not made from text but from images, which means that you can't change the fonts)

To open a subtitles file, use the Advanced Open dialog box (Menu File, Open file). Select your file by clicking on the *Browse* button. Then, check the *Subtitle options* checkbox and click on the Settings button.

 Select a subtitles file under Windows or Linux

You can then select the subtitles file by clicking the *Browse* button. You can also set a few options like character encoding, alignment and size. The delay option allows you to delay the subtitles against the video if they are not in sync. If they are not at the same speed, you might also want to adjust the subtitles framerate.

Note: For Vobsub subtitles, you need to select the **.idx** file, not the **.sub** file. Encoding, alignment and size won't have any effect for Vobsub subtitles.

Font can be changed as explained in the previous section.

#### Video and audio filters

VLC includes a system of *filters* that allow you to modify the audio and video.

##### Deinterlacement and Post Processing

VLC is able to deinterlace a video stream using different deinterlacement methods. Deinterlacement can be enabled in the *Video* menu, *Deinterlacement* menu item. The *Blend* methods gives the best results in most cases. The *discard*method is a less resource consuming alternative.

On some particular streams (MPEG 4, DIVX, XVID, Sorenson, etc.), some additional image filtering can be applied to the video before display, improving its quality in some cases. This can be enabled in the *Video* menu, *Post processing* menu item. Different levels of post processing can be chosen here. A higher level means more filtering.

##### Video filters

VLC features several filters able to change the video (distortion, brightness adjustment, motion blurring, etc.).

With the wxWidgets interface, filters can be easily enabled using the Extended GUI. In the Video tab, simply select the filters to enable. Image settings can be easily adjusted.

 Video filters selection in the wxWidgets interface

You can enable these filters through the *Extended Controls panel* on Mac OS X. Click on the triangle next to *Video filters* to select your filters or expand the *Adjust Image* section to change the contrast, hue, etc.

 Video filters selection in the Mac OS X interface

For better control, you need to go to the preferences. To select the filters to be enabled, go to *Video*, then to *Filters*. In the "video filter module" box, enter the names of the filters to enable, separated by semicommas. Filters will be applied in the selected order. Valid names are "clone", "wall", "transform", "adjust", "crop", "deinterlace", "distort", "motionblur" and "logo".

If you want to tune the behavior of these filters, go to *Video, Filters, \[your filter\]*. For each filter, you will find a short description and the options.

##### Audio filters

###### Equalizer

VLC features a 10-band graphical equalizer. You can display it by activating the advanced GUI on wxWidgets or by clicking the *Equalizer* button on the MacOS X interface.

 The equalizer in the wxWidgets interface

 The equalizer in the MacOS X interface

Presets are available in the Audio menu in wxWidgets, or in the Equalizer window in the MacOS X interface.

###### Other audio filters

At the moment, VLC features two other audio filters: a volume normalizer and a filter providing sound spatialization with a headphone. They can be enabled in the Audio tab of the extended GUI for the wxWidgets interface and in the Audio section of the Extended Controls panel of the Mac OS X interface.

For better control, you need to go to the preferences. To select the filters to be enabled, go to *Audio*, then to *Filters*. In the "audio filters" box, enter the names of the filters to enable, separated by commas. Valid names are "equalizer", "normvol" and "headphone".

If you want to tune the behavior of these filters, go to *Audio, Filters, \[your filter\]*. The equalizer and headphone filters can be tuned.

#### Snapshots (aka, screenshots)

There are two ways to take snapshots (i.e., screenshots or frame grabs) with VLC:

1.  Go to Video -\> Snapshot, or
2.  Press the snapshot hotkey
    -   Windows / Linux / Unix: Ctrl-Alt-s
    -   Mac OS X: Command-Alt-s

When a snapshot is taken, it will briefly preview as a thumbnail with its filename and then fade away.

To change the hotkey, go to Preferences -\> Interface -\> Hotkeys settings. Check Advanced options, and set Take video snapshot.

##### Snapshot location, format and name

The snapshot location depends upon your operating system:

-   Windows: My Documents\\My Pictures\
-   Linux / Unix: \$(HOME)/.vlc/
-   Mac OS X: Desktop/

The default format for snapshots is PNG, but this may be changed to JPEG. Also, the default name for snapshots is *vlcsnap-* followed by a timestamp that is *not* the time of the frame in the video you're viewing.

The location, format and name of snapshots may be changed in the Preferences. Also, you may substitute other text for *vlcsnap-* in the *Video snapshot file prefix* and you may choose to have snapshots numbered sequentially (i.e., 000001, 000002, 000003, and so on) instead of with a timestamp. As of version 0.9.0, you may even use variables in the text used for the filename. For example, *\$T* (must be upper case) will insert the video's time code into the file name. If you were to change the prefix to *Friends-\$T-* while watching a DVD of *Friends*, then the snapshot filenames would look something like this: Friends-00_05_21-00004.png . This indicates a snapshot taken at 5 minutes and 21 seconds into the video; and it was the number 00004 snapshot of the day.

For a full list of variables, please see Documentation:Play HowTo/Format String.

#### Hotkeys

Most of VLC functions are accessible using hotkeys.

The list of the available hotkeys and their functions can be retrieved and altered in the preferences panel of the player. In the wxWidgets interface, preferences are available in the "Settings" menu, "Preferences" menu item. In the MacOS X interface, open the "VLC" menu, and select "Preferences". Select the "Hot keys" panel in the dialog.

As of version 0.9, a list of hotkeys is presented in a drop-down window. To change one, double-click its name to select it. Then, press the new key that will trigger the specified action. Modifier keys (such as Control/Command and Alt) may also be used.

In earlier versions, several boxes give the list of modifiers for the hotkey. To trigger an action using a hotkey, you need to press simultaneously the keys corresponding to the different selected modifiers as well as the key set in the dropdown.

To change the binding of a hotkey, select or deselect boxes corresponding to the different modifiers, and change the key by using the drop-down menu. Select the *Save* button to apply the changes.

 The Hotkeys Panel - wxWidgets interface

 The Hotkeys Panel - MacOS X interface

#### Basic troubleshooting

##### File does not play, only sound or only video

Maybe the file you are trying to read is not fully supported. VLC does not use the codec packs (the software that decodes video signals) you might have installed. It comes with its own codecs. If there is no open-source decoder for the format you are trying to read, it won't be supported. (There is an exception, under Windows, for codecs that use the DirectShow framework.)

To find out, open the Messages Window (View menu) and restart your stream. Look for error messages (red messages)

 The wxWidgets messages window

In this example, the file contains a IV41 video stream, a codec that is not supported by VLC.

You may of course have other messages. If you post to a VideoLAN mailing list or in the forum, please include such a log. It is very valuable in troubleshooting.

##### Weird VLC behavior and crashes

A very common thing is a corrupted VLC preferences file. Don't hesitate to delete it if problems appear suddenly. You will find in the FAQ details on [how to delete your preferences file](http://www.videolan.org/doc/faq/en/index.html#id2470084).

##### Computer crashes / Video is corrupted

Another common problem is buggy video drivers. Try upgrading them from the website of your video card's manufacturer.

Also, you can try disabling Overlay (Preferences/General/Video, untick "Overlay video output")

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use 0.9 {#play-howto-basic-use-0-9}

#### Overview of the VideoLAN project

VideoLAN was a complete software solution for video streaming and playback, developed by students of the [Ecole Centrale Paris](http://www.ecp.fr) and developers from all over the world, under the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) (GPL).

Originally VideoLAN was designed to stream MPEG videos on high-bandwidth networks, but VideoLAN's main software, VLC media player, has evolved to become a full-featured, cross-platform media player.

Now the Non-Profit Organisation developing and offering the VLC media player is called: VideoLAN Organisation

More details about the project can be found on the [VideoLAN Web site](https://www.videolan.org/).

#### VLC Media Player

VLC 2.0 default interface, Windows

Originally called *VideoLAN Client*, VLC media player is VideoLAN's main software product.

VLC media player works on many platforms: Linux, Windows, macOS, BeOS, BSD, Solaris, Android, iOS, QNX and many more... It supports the following video and audio formats: MPEG-1, MPEG-2, MPEG-4/DivX, h264, webm, mkv, DVDs, VCDs, Audio CDs, wmv and wma.

It can also play from external sources:

-   Satellite.
-   Cable.
-   Digital TV cards (DVB-S, DVB-T).
-   Several types of network streams: UDP/RTP Unicast, UDP/RTP Multicast, HTTP, RTSP, MMS, etc.
-   Acquisition or encoding cards.
-   Webcams and other devices.

VLC can also be used as a streaming server. This feature is described in the [Streaming HowTo](#streaming-howto).

This guide describes all the playback (client) aspects of VLC media player.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use 0.9 / Audio {#play-howto-basic-use-0-9-audio}

VLC can play several audio formats: *.asf, .avi, .divx, .dv, .mxf, .ogg, .gm, .ps, .ts, .vob,* and *.wmv*. It can convert audio tracks and use several visualizations.

**Note:** The commands in the **Audio** menu are only enabled when an audio file is being played.

#### Playing an audio track

To play a track:

1.  Select *Open File* in the *Media* menu.
2.  Select an audio file and click on the  *Play* button. The selected track is played.

#### Enabling and disabling audio tracks

-   To disable a track, select the *Disable* option in the *Audio Track* from the *Audio* menu. The selected track will then stop.
-   To enable the track again, select the designated *Track* option in the *Audio Track* from the *Audio* menu. The selected track will then play.

#### Recording Audio

To record audio you need the record button () to be visible. The record button is hidden by default. You can display using one of these methods:

-   Select Advanced Controls in the View menu. The Advanced toolbar is displayed on top of the standard toolbar. The Advanced toolbar contains the Record button.
-   Select Customize interface in the Tools menu and add the record button to the Line 2 of buttons (which is the line shown by default).

Once the Record button is visible, click it to start recording.

The recording from a shoutcast stream is stored somewhere in your files under a name like 0 (e.g.: 1, when recording from [Radio CAFF](http://radiocaff.com.ar/) (or more precisely from the underlying [WinAmp stream](http://panel7.serverhostingcenter.com/tunein.php/radiocaff/playlist.pls)). Under my german Windows XP it was stored under "Eigene Dateien/Eigene Music" so I guess that you find it in an english Windows under "My Documents/My Music/", I don't know where it will be stored under Linux or any other OS (updates are welcome).

You can automagically cut the stream into tracks by relaying the stream through [Streamripper](http://streamripper.sourceforge.net), i.e. by directing StreamRipper to the ShoutCast stream and directing VLC to the relaying port of StreamRipper (default http://localhost:8000).

#### Audio Device

This option helps you to listen to audio files in two modes: stereo and mono.

1.  To listen to an audio track in either the Stereo or Mono mode, select *Open File or Open Disc* from the *Media* menu. The Open dialog box is displayed.
2.  Select an audio file and click on the  *Play* button. The selected track is played.
3.  Select *Mono* in *Audio Device* from the *Audio* menu if you want to listen to the audio track in the Mono mode.

Mono refers to monaural sound that uses a single channel for sound reproduction.

1.  Select *Stereo* in *Audio Device* from the *Audio* menu if you want to listen to the audio track in the Stereo mode.

Stereo refers to sound that uses two channels for sound reproduction or stereophonic sound.

#### Audio Channels

In audio, a channel refers to a stream of audio that is to be played by one speaker. For example, stereo audio, consists of two channels. This option is useful for codecs that don’t have support for more than 2 channels.

Select a channel type in *Audio Channels* from the *Audio* menu. VLC media player provides four audio channels and they are:

1.  *Stereo* – Refers to the reproduction of the sound in two or more independent audio channels using more than one speaker. If you use this option, you would feel as though the sound is played from all the directions. You can observe this in a regular home theatre with 5.1 or 6.1 speakers.
2.  *Left* – You can observe this in a regular audio player with 2.1 speakers. If you select the **Left** option, the music is played only in the left speaker. The speaker on your right is automatically switched OFF.
3.  *Right* - If you select the **Right** option, the music is played only in the speaker on your right side. The speaker on your left is automatically switched OFF.
4.  *Reverse Stereo* – There are several applications that are used to reverse the stereo whereas VLC has an in-built feature to reverse the stereo. This option is useful if you want the audio to play in tandem with the video. You can use the **Reverse Stereo** option if you want to deliberately change the audio output.

Imagine that you are watching a video. In the video, a person walks on the left side but the sound is produced on the right speaker. You can correct this by selecting the *Reverse Stereo* option in VLC. Select the *Reverse Stereo* option and play the same scene in the video and observe the difference.

You can observe this with 2.1, 5.1, 6.1 and 8.1 speakers.

#### Visualize Audio

Visualizations display splashes of colour and geometric shapes and generate animated imagery based on a piece of music.

The different visual effects available are *Spectrometer, Scope, Spectrum, VU Meter and Goom*. This menu item can also be used to disable a visualization.

1.  Select an option under the *Visualizations* option from the *Audio* menu to view the effects. The selected visualization is then played.
2.  To disable visualizations, select *Disable* under *Visualizations* from the *Audio* menu. The visualization is then disabled.

Spectrum visualization on VLC:

#### Maximum VLC Volume

To change the maximum volume in % that VLC should use, go to **Tools** → **Preferences** (select **All** at bottom left corner) → **Interface** → **Main interfaces** → **Qt** → **Maximum volume displayed**.

Save it and restart VLC.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use 0.9 / Basic troubleshooting {#play-howto-basic-use-0-9-basic-troubleshooting}

**Languages:English** • français

#### VLC Support Guide: Solve your VLC issues right now!

The **V**LC **S**upport **G**uide is an informal, step-by-step guide for troubleshooting most common issues with VLC.

It complements the VLC media player Documentation.

**So what's your problem?**

##### Installation Issue

VLC won't install. Go!

##### Startup Issue

VLC won't start up. Go!

##### Audio Playback Issue

The audio or the sounds are wrong. Go!

##### Video Playback Issue

The video is messed up. Go!

##### Subtitle Display Issue

The subtitles aren't working properly. Go!

##### Usage Issue

I have difficulty using VLC. Go!

##### Interface Issue

I want to change my interface. Go!

##### Uninstallation Issue

VLC won't uninstall (why are you uninstalling it anyway?). Go!

#### Get Help

If this troubleshooter does not resolve your problems or answer your questions, some other resources which you can use include:

-   Frequently asked questions
-   Frequently asked questions about VLC on Windows
-   Frequently asked questions about VLC on macOS
-   Frequently asked questions about VLC on Linux
-   The [VideoLAN support forum](https://forum.videolan.org/)
-   The VideoLAN IRC channel.
-   VLC documentation

This page is part of the informal VLC Support Guide.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use 0.9 / Hotkeys {#play-howto-basic-use-0-9-hotkeys}

Most of VLC functions are accessible using hotkeys.

The list of the available hotkeys and their functions can be retrieved and altered in the *Preferences* panel of the player. In the Windows and Linux interface, *Preferences* are available in the "Tools" tab as the "Preferences" menu item. In the MacOS X interface, open the "VLC" menu, and select "Preferences". Select the "Hot keys" panel in the dialog.

As of version 0.9, a list of hotkeys are presented in a drop-down window. To change one, double-click its name to select it. Then, press the new key that will trigger the specified action. Modifier keys (such as Control/Command and Alt) may also be used. In the 1.x version you can also filter hotkeys with a search filter.

In earlier versions, several boxes gave the list of modifiers for the hotkey. To trigger an action using a hotkey, you need to press simultaneously the keys corresponding to the different selected modifiers as well as the key set in the dropdown.

To change the binding of a hotkey, select or deselect boxes corresponding to the different modifiers, and change the key by using the drop-down menu. Select the *Save* button to apply the changes.

The Hotkeys Panel - MacOS X interface**FIXME - needs verifying for 0.9**

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use 0.9 / Interface {#play-howto-basic-use-0-9-interface}

#### General Interface Description

VLC has several interfaces:

-   A cross-platform interface for Windows and GNU/Linux, which is called Qt.
-   A native Mac OS X interface.
-   An interface that supports skins for both Windows and GNU/Linux.

The operation of VLC is essentially the same in all the interfaces.

##### Windows and GNU/Linux (Qt)

The screenshot below shows the default interface in VLC 2.0. More features can be displayed by selecting them in the *View* menu.

See also VLC Interface 2.0 on Windows 7

##### Mac OS X

This screenshot shows the default interface that VLC had on Mac OS X until version 1.1:

Since version 2.0 the interface has been redesigned. See OSX 2.0 interface.

#### Starting VLC Media Player in Windows

In Windows XP: Click **Start** -\> **Programs** -\> **VideoLAN** -\> **VLC media player**.

In Windows 7: Click **Start** -\> **All Programs** -\> **VideoLAN** -\> **VLC media player**.

VLC is shown on the screen and a small icon  is shown in the system tray.

#### Stopping VLC Media Player

There are three ways to quit VLC:

-   Right click the VLC icon () in the tray and select **Quit** (*Alt-F4*).
-   Click the **Close** button in the main interface of the application.
-   In the **Media** menu, select **Quit** (*Ctrl-Q*).

#### Notification Area Icon

Clicking this icon shows or hides the VLC interface. Hiding VLC does not exit the application. VLC keeps running in the background when it is hidden. Right clicking the icon in the notification area shows a menu with basic operations, such as opening, playing, stopping, or changing a media file.

#### Main Interface

The main interface has the following areas:

-   **Menu bar**.
-   **Track slider** - The track slider is below the menu bar. It shows the playing progress of the media file. You can drag the track slider left to rewind or right to forward the track being played. When a video file is played, the video is shown between the menu bar and the track slider.
    **Note: When a media file is streamed, the track slider does not move because VLC cannot know the total duration.**
-   **Control Buttons** - The buttons below the track slider cover all the basic playback features.

Click here to view an explanation of every menu item.

#### Opening media

See Documentation:Play HowTo/Basic Use 0.9/Opening modes

#### Streaming Media Files

Streaming is a method of delivering audio or video content across a network without the need to download the media file before it is played. You can view or listen to the content as it arrives. It has the advantage that you don't need to wait for large media files to finish downloading before playing them.

VideoLan is designed to stream MPEG videos on high bandwidth networks. VLC can be used as a server to stream MPEG-1, MPEG-2 and MPEG-4 files, DVDs and live videos on the network in unicast or multicast. Unicast is a process where media files are sent to a single system through the network. Multicast is a process where media files are sent to multiple systems through the network.

VLC is also used as a client to receive, decode and display MPEG streams. MPEG-1, MPEG-2 and MPEG-4 streams received from the network or an external device can be sent to one machine or a group of machines.

**To stream a file**:

1.  From the **Media** menu, select **Open Network Stream**. The *Open Media* dialog box loads with the *Network* tab selected.
2.  In the **Please enter a network URL** text box, Type the network URL.
3.  Click **Play**.

Note: When VLC plays a stream, the track slider shows the progress of the playback.

For more information, refer to Documentation:Streaming HowTo/Receive and Save a Stream

#### Converting and Saving a Media File Format

VLC can convert media files from one format to another.

**To convert a media file**:

1.  From the **Media** menu, select **Convert/Save**. The *Open media* dialog window appears.
2.  Click **Add...**. A file selection dialog window appears.
3.  Select the file you want to convert and click **Open**. The *Convert* dialog window appears.
4.  In the **Destination file** text box, indicate the path and file name where you want to store the converted file.
5.  From the **Profile** drop-down, select a conversion profile.
6.  Click **Start**.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use 0.9 / Opening modes {#play-howto-basic-use-0-9-opening-modes}

#### Opening a File

The **Media** menu can be used to open a file. VLC offers a range of options to open media files. See the table below to see the available options. When you open a file, the file is played according to the default play options.

-   Option: **Shortcut KeyDescription**

-   **Open File**: Ctrl+O Use this option to play a single media file from a specified location on the hard disk.

-   **Advanced Open File**:  

    In addition to opening a file from a hard disk, you can open files from a disc, from any computer on the network or directly from a capturing device.

    You can also open a subtitles file associated with the selected media file.

    You can also set a few playing options. Refer to Advanced File Open.

-   **Open Folder**: Ctrl+F

    Use this option to play all the files in a certain folder.

-   **Open Disc**: Ctrl+D

    Use this option to play files from a disc. Based on the type of disc you select, you can have a few more playing options. Refer to Opening a file from a disc.

-   **Open Network**: Ctrl+N

    Use this option to open a file present on any system on the network to which you are currently connected.

    You can also set a few playing options. Refer to Opening a file on the network

-   **Open Capture Device**: Ctrl+C

    Use this option to open a file directly from a capturing device which is currently connected to your system.

    You can also set a few playing options. Refer to Opening a file from the capturing device.

-   **Convert / Save**: Ctrl+R

    Use this option to convert a media file from one format to another.

    Refer to Converting a media file into another format.

-   **Streaming**: Ctrl+S

    Use this option to stream a recorded media file.

    Refer to Streaming a media file.

You can open audio or video files. A file can be opened in two ways.

1.  From the **Media** menu select the **Open File** option or the **Advanced Open File** option.
2.  Select the **File** tab in the Open dialog box.
3.  Enter file name in the **File names** box or browse and select a file.
4.  Select a format from the **Filter** list. The supported formats are *.a52, .aac, .ac3, .dts, .p3, .ogg, .oma, .spx, .wav, .wma* and *.wm*.
5.  Check the **Use a subtitle file** option to select a subtitles file to be viewed with the media file.
    From the **Alignment** list, select an option to align selected subtitle. The available options are *Left, Right* and *Center*.
    From the **Size** list, select a font size for the selected subtitle. The available options are *Small, Smaller, Normal, Large* and *Larger*.
6.  Click **Open**. VLC starts playing the selected file with the default options.

#### Advanced File Open

To open a file

1.  Select the **Advanced Open File** option from the **Media** menu.
2.  The Open file dialog box is displayed. There are four tabs such as **File**, **Disc**, **Network** and **Capture Device**.

Refer to the following sections for more details:

Opening a media file with advanced options
Opening a folder
Opening a media disc
Opening a media network
Opening a capture device

#### Additional Playing Options

When you select the **File** tab after selecting the **Advanced Open File** menu, apart from selecting a file you have the following choices:

**Caching:** When you specify caching value for a media file, the stream is still rendered by VLC media player at the specified data rate, but the client system buffers a much larger portion of the content before rendering it. This allows the client to handle variable network conditions without a perceptible impact on the playback quality of either on-demand or broadcast content. Specify a caching value so that the file is played smoother. The default value is 300 milliseconds.

**Customizing:** File names from different locations can be added directly into the **Customize** box without having to browse the folders.

**Synchronous play:** Play another file in synch with the selected file.

**Start time:** Not play the file from the beginning. If start time is specified as 120 seconds, the file is played after skipping the content of the first two minutes. (Specify time in seconds; not minutes)

#### Opening a Folder

You can select a folder to play all media files one after the other in that folder.

1.  Select the **Open Folder** option. The Browse for Folders dialog box is displayed.
2.  Browse and select the folder.
3.  Click on **OK**.

All the files present in the selected folder are played in the alphabetical order, one after another, without expecting any action from you.

#### Opening a Disc

You can open media files from a disc. In VLC, you can play Audio CDs, SVCD/VCDs, and DVDs. You can open a file from a disc in two ways.

1.  Select the **Open Disc** option from the **Media** menu.
    *Or*
    Select the **Advanced Open File** option from the **Media** menu.

2.  Select the **Disc** tab.

3.  Check the type of disc connected to the system. The options are **Audio CD**, **DVD**, and **SVCD/VCD**.

4.  In the **Disc Device** box, by default, the path for the disc is displayed. You may select a different path using the **Browse** button.
    Click on the Eject  button. The disk drive opens automatically and you can check if the drive is empty or if the correct disc is in the drive.

5.  Based on the selected disc type, specify the following options:

6.  -   DVD

    -   -   Some original DVDs may have complex, proprietary menu options and VLC may not handle all the options well. If you check the **No DVD menus** option, VLC reads the raw video files directly into the film regardless of the options present while creating the original DVD. Check this option if you want to listen to or view the basic version without availing the menus present in the DVD.
        -   When a DVD is played, the entire disc need not be played. To specify the part to be played, specify the **Title** and **Chapter number** in the **Starting Position** box.
        -   Under the Audio and Subtitles group, select the **Subtitles track** and **Audio track**.

    -   Audio CD

    -   SVCD/VCD

7.  Check **Show more options** to see more play options. Refer to Additional playing options.

##### Playing more than one media file

VLC has an option to play two media files synchronously.

1.  From the **Media** menu, select the **Advanced Open File** option.
2.  Select a media file.
3.  Check the **Show more options** option. The screen expands to show more options.
4.  Check the **Play another media synchronously** option. The **Extra media** box and a **Browse** button are displayed.
5.  In the **Extra media** box, enter the name of another media file with complete path or use the **Browse** button to select the media file.
6.  You can change the **Caching** value for the media file being played.
7.  Select the time at which the media file should start from the **Start Time** list.
8.  You can see the selection you made in the **Customize** box.
9.  Click on the **Play** button.

You can use the **Show more options** to watch a video while listening to an audio file or listen to two audio files played synchronously (one audio track can have the instrumental part and the other can have the corresponding voice).

#### Opening a Network

You can open a network and stream media from the selected network to the specified hosts. When you open a network you specify the network to be used for streaming media content.

1.  From the **Media** menu select **Open Network** or **Advanced Open File** option and then select the **Network** tab.

2.  Select a protocol from the **Protocol** list.

3.  Select a protocol suitable to your content.

4.  In the **Address** box, enter the address of the system from which the media is going to be streamed.

5.  In the **Port** box, enter the port number from which streaming is done.

6.  When UDP is selected, the **Allow Timeshifting** option is enabled.

7.  Enter a URL in the **Address** box.

8.  Click on the  before the **Play** button and select **Stream** from the popup menu.

9.  In the Stream Output dialog box, specify the media file to be streamed and the address to which the streaming should be done.
    In the Stream Output dialog, you can specify further options. Refer to Specifying the Streaming Options.

10. Click on the **Stream** button.

11. 1.  Select **Streaming** from the **Media** menu.
    2.  Click on the  icon next to the **Play** button and select **Stream** from the popup menu.

12. 1.  Check the **Play locally** option to play the file while it is being streamed.

    2.  Check the **File** option to specify a path to save the converted file or click on the **Browse** button. The Save File dialog box is displayed. Select a container format from the **Save As Type** list.

    3.  -   Format: **Description**
        -   **.ps**: Refers to MPEG program stream. Stores M-PEG 2 video muxed with other streams.
        -   **.ts**: Refers to MPEG transport stream. Used for streaming video through a network or by a satellite.
        -   **.mpg**: Refers to a family of standards used for coding audio and visual information.
        -   **.ogg**: Refers to professional grade media product. Ogg Vorbis encodes audio and Ogg Theora encodes video.
        -   **.asf**: Stores Windows Media Audio and Windows Media Video. ASF is designed to be used over audio and video information and is specially designed to run over networks.
        -   **.mp4**: M-PEG 4 audio and video. Provides compression for web, voice and broadcast television applications.
        -   **.mov**: Refers to the QuickTime media format. Used to store audio and video.

    4.  Select a file or enter the file name in the **File** name box.

    5.  Click on **Save** to save the media file in the selected container format.

    6.  Check the **Dump Raw Output** box to save the input stream as it is read by VLC, without any processing. If this option is selected, all other options are disabled.

    7.  Select HTTP to stream media files using the HTTP streaming method. Specify the **Address** and **Port**.

    8.  Select the **MMSH** access method to stream media files to the Microsoft Windows Media Player. The **Address** and **Port** options are enabled. Specify the **Address** and **Port**.
        MMS is a proprietary digital media streaming protocol developed by Microsoft. MMSH is MMS over HTTP.

    9.  Select **RTP** to stream the media using the RTP method. The Prefer UDP over RTP, Address, Port, Audio Port and Video Port options are enabled.
        RTP refers to the Real-Time Transfer Protocol. Like UDP, RTP can use both unicast and multicast addresses. RTP or UDP is extensively used for streaming live audio and video.

    10. Specify the **Address**, **Port**, **Audio Port** and **Video Port**.

    11. Select the **Prefer UDP over RTP** option.

    12. Select **IceCast** to distribute live audio and video over the Internet in real time.
        -   Enter the **Address** and **Port** details.
        -   Enter the login name and password in the **Login:pass:** box.
        -   Enter the name of the **Mount Point** where the current listener should be redirected to.

        An IceCast mount point refers to a connector between an IceCast source stream and IceCast listeners.

    13. Select a profile from the **Profile** list. The available profiles are *Custom, Ogg/Vorbis, MPEG-2, MP3, MPEG-4 audio AAC, MPEG-4/DivX, H264, IPod (MP4, aac), Xbox, Windows (wmv/asf),* and *PSP*.

    14. Choose the encoder format from the **Profiles** or customise it.

    15. Customise the other options by selecting the Encapsulation, Video codec, Audio codec and Subtitles tabs.

    16. -   Select the required codec from the **Codec** list. The available video codecs are *MPEG-1, MPEG-2, MPEG-4, DIVX 1, DIVX 2, DIVX 3, H-263, H-264, WMV1, WMV2, MPEG,* and *Theora*.
        -   Specify an average bitrate in the **Bitrate** (kb/s) box.
        -   Select a scale from the Scale list. The values are *1, 0.25, 0.5, 0.75, 1.25, 1.5, 1.75,* and *2*.

    17. -   Select an audio codec from the **Codec** list. The available audio codecs are *Vorbis, MPEG Audio, MP3, MPEG4 Audio (AAC), A52/AC-3, Flac, Speex, WAV* and *WMA*.
        -   Specify an average bit rate in the **Bitrate** (kb/s) box.
        -   Select a channel from the **Channels** list. In audio, a channel refers to a stream of audio that is to be played by one speaker. For example, stereo audio, consists of two channels.

    18. -   Check the **Subtitles** checkbox and select a subtitle from the **Subtitle** list.
        -   Check the **Overlay subtitles on the video** option to render subtitles directly on the video, while transcoding it.

    19. Click on the Stream button. The selected file is streamed to the selected locations.

    20. -   Option: **Shortcut KeyDescription**

        -   **Enqueue**: Alt + E Adds media files to the playlist but doesn't play it until you click **Play**.

        -   **Play**: Alt + P Adds media files to the playlist and plays the media.

        -   **Stream**: Alt + S Adds media files to the playlist and streams it on the network.

        -   **Convert**: Alt + C Adds media files to the playlist.

            Converts a media file into the selected format.

    21. 1.  Select **Open Capture Device** from the **Media** menu. The Open dialog box is displayed with the **Capture Device** tab selected.

        2.  1.  Click on the **Configure** button for Video. The Properties dialog box is displayed with two tabs, **Device Settings** and **Advanced**.

            2.  If the device name does not appear in the list, click on the **Refresh** button. The device name appears in a list next to the **Configure** button.

            3.  -   **Brightness:** Move the slider till you get the desired brightness for the video capture. The default value is 5000.
                -   **Contrast:** Refers to the difference in visual properties that makes an object distinguishable from other objects and the background. Move the slider till you get the desired contrast. The default value is 5000.
                -   **Saturation:** Refers to the difference of a color against its own brightness. Move the slider to get the desired effect. The default value is 5000.
                -   **Sharpness:** Refers to the clarity of a video. Move the slider till you get the desired sharpness for the video capture. The default value is 6000.
                -   **White Balance:** Refers to colour balance. This option helps to make white actually white and makes skin tones look more natural. Uncheck the **Auto** option and Move the slider to get the desired effect.
                -   **Backlight Comp:** Refers to the ability of a camera to compensate in cases where a subject with a large amount of background light would otherwise be obscured by excessive light. The default value is 0. Move the slider to get the desired effect.

            4.  -   **Exposure** - Refers to the amount of light allowed to fall on a selected media file while capturing images. There are occasions when you may have to manually adjust the exposure on your camera. Exposure is measured in seconds.
                    For example, you have to take a shot of a person from a certain angle, and there is bright light behind the person. In such case, aim your camera on the person and adjust the exposure value by moving the slider. The specified value remains unchanged even after closing the VLC application.
                -   **Gain** - This option allows increasing or decreasing the brightness of the video being captured.
                    When **Automatic Gain Control** is selected, the values you specified are taken as the default values for Exposure and Gain.
                    Uncheck **Automatic Gain Control** to change the values of **Exposure** and **Gain** by moving the sliders.

            5.  -   **Mirror Horizontal** – If you select this option, the video clip is flipped horizontally. You can see a mirror view of the captured picture.
                -   **Mirror Vertical** - If you select this option, the video clip is flipped upside down.

            6.  -   **Low Light Boost** – If you check this option, the exposure time of the camera increases in poor light conditions.
                -   **Color Boost** – If you check this option, the colors of the video being captured are boosted.

            7.  -   **Loudness** – Refers to volume of the audio. Adjust the volume by moving the slider.
                -   **Mono** – Refers to an amplifier connection. Adjust the volume by moving the slider.

            8.  Click on the **Advanced options** button to specify the following properties:

            9.  -   **Caching value in ms** – Refers to the caching value for DirectShow streams. Enter or select a value.

                -   **Video device name** – Refers to the name of the video device that is used by DirectShow plugin. If you do not specify a device, the default device is used.

                -   **Audio device name** – Refers to the name of the audio device that is used by DirectShow plugin. If you do not specify a device, the default device is used.

                -   **Video size** – Refers to the size of the video that is displayed by the DirectShow plugin. The size of video is measured in pixels. If you do not specify the size, the default size is used.

                -   **Video input chroma format** - Chroma refers to the way colors are encoded. Enter a specific chroma format. The default value is 1420.

                -   **Video input frame rate** – Enter a specific frame rate. The default value is 0.

                -   **Device properties** – Check this option to view the properties dialog of the selected device before starting the stream.

                -   **Tuner properties** – Using this option you can set channels. A tuner converts signals into picture and sounds. Select this option to view the tuner properties (channel selection) dialog box.

                -   -   **Tuner TV channel** – Refers to a tuner for setting TV channels. The default is 0. The default channel is used to capture the media.

                    -   **Tuner country code** – This option helps to establish the current channel-to-frequency mapping. The default is 0.

                    -   **Tuner input type** – Select the tuner input type. Available values are cable and antenna.

                    -   **Video input pin** – Select a video input source. Available values are Composite, S-video, and Tuner. These settings are hardware-specific. -1 means that settings will not be changed.

                    -   **Audio input pin** – This option is used to capture audio using a specific audio input pin. These settings are hardware-specific. Select a number from the Audio input pin list.

                    -   **AM Tuner mode** – This option is used to select a AM (amplitude modulation). The following are the tuner modes:

                    -   -   Value: **Mode**
                        -   **0**: Default
                        -   **1**: TV
                        -   **2**: AM Radio
                        -   **3**: FM Radio
                        -   **4**: DSS

                    -   Number of audio channels – Select an audio input format with the given number of audio channels. If the channels are unavailable, select 0.

                    -   Audio sample rate – This option is used to set the sample rate. If the rates are unavailable, select 0.

                    -   Audio bits per sample – Select audio input format with the given bits or sample. If the audio bits are unavailable, select 0.

            10. Select the **Convert** option to select the encoding formats and click on the **Save** button. Refer to Converting and Saving a media file format.

            11. Click on the **Play** button. The capturing of the media starts.

        3.  -   **DVB-S** - Is an abbreviation for Digital Video Broadcasting - Satellite. It is the Digital Video Broadcasting forward error coding and modulation standard for satellite television. This is used via satellites.
            -   **DVB-C** – Is an abbreviation for Digital Video Broadcasting – Cable. It is the DVB European consortium standard for the broadcast transmission of digital television over cable. This system transmits MPEG-2 or MPEG-4 audio and video streams using a QAM modulation.
            -   **DVB-T** - Is an abbreviation for Digital Video Broadcasting – Terrestrial. It is the DVB European-based consortium standard for the broadcast transmission of digital terrestrial television. This system transmits compressed audio, video and other data in the MPEG format using the COFDM modulation.

        4.  1.  Select **Open Capture Device** from the **Media** menu. The Open dialog box is displayed.

            2.  Select the **Open Capture Device** tab.

            3.  Select **DVB DirectShow** from the **Capture Mode** list.

            4.  Select **DVB-S** from **DVB Type** under the **Card Selection** group. In the **Options** group, specify the following

            5.  Select **Transponder/multiplex frequency** to set the transponder frequency. A transponder is a device that receives, amplifies and retransmits a signal on a different frequency.

            6.  Select **Transponder symbol rate** to set the transponder symbol rate.

            7.  Click on the **Advanced options** button to specify the following parameters:

            8.  -   **Caching value in ms** – Refers to caching value for the DirectShow stream. Enter a value in milliseconds.
                -   **Transponder / multiplex frequency** - A transponder is a device that receives, amplifies and retransmits a signal on a different frequency. Select a frequency.
                -   **Inversion Mode** - *Description to be added*
                -   **Satellite polarization** – Polarization is a method of giving transmission signals a specific direction. The signals transmitted by a satellite can be polarized in four ways and they are: *Horizontal, Vertical, Circular Left* and *Circular Right*. Select an option.
                -   **Network identifier** – Refers to a unique ID used to identify a network. Select a number from the **Network identifier** list.
                -   **Satellite Azhimuth** – Azhimuth is an angular measurement made in the horizontal plane. Enter a value.
                -   Satellite Elevation– This option defines the angle between the Earth and the position of a satellite. Enter a value.
                -   **Satellite Longitude** – Refers to the satellite longitude in 10ths of degree. Enter a value.
                -   **Antenna lnb_lof1** – Refers to low band local Osc Freq in kHz. Enter a value in kHz.
                -   **Antenna lnb_lof2** - Refers to high band local Osc Freq in kHz. Enter a value in kHz.
                -   **Antenna lnb_slof** – Refers to low noise block switch freq in kHz. Enter a value in kHz.
                -   **Transponder FEC** – Refers to the forward error correction mode. Enter a value in kHz.
                -   **Transponder symbol rate in kHz** *Description to be added*
                -   **Modulation Type** – Refers to the QAM constellation points. The available values are *16, 32, 64, 126,* and *256*.
                -   **Terrestrial high priority stream code rate (FEC)** – Refers to the high priority FEC Rate. The available values are Undefined, *1/2, 2/3, 3/4, 5/6* and *7/8*.
                -   **Terrestrial low priority stream code rate (FEC)** – Refers to the low priority FEC Rate. The available values are *Undefined, 1/2, 2/3, 3/4, 5/6,* and *7/8*.
                -   **Terrestrial bandwidth** - *Description to be added*
                -   **Terrestrial guard interval** – Refers to a parameter that is used in encoding and modulation. Select an interval from the list.
                -   **Terrestrial transmission mode** - *Description to be added*
                -   **Terrestrial hierarchy mode** - *Description to be added*

        5.  1.  Select **Open Capture Device** from the **Media** menu. The Open dialog box is displayed.
            2.  Select the **Open Capture Device** tab.
            3.  Select **DVB DirectShow** from the **Capture Mode** list.
            4.  Select **DVB-C** from **DVB Type** under the **Card Selection** group.
            5.  Select **Transponder/multiplex frequency** to set the transponder frequency.
            6.  Select **Transponder symbol rate** to set the transponder symbol rate.
            7.  Select an extra media if you want some background music using **Show more options**. Refer to \\[#Playing_more_than_one_media_filePlaying more than one media file.
            8.  Select **Convert** to select the encoding formats and click on the **Save** button. Refer to Converting and Saving a Media File Format.
            9.  Click on the **Play** button to play the media.
            10. Click on the **Cancel** button to exit the screen.

        6.  1.  Select **Open Capture Device** from the **Media** menu. The Open dialog box is displayed.
            2.  Select the **Open Capture Device** tab.
            3.  Select **DVB DirectShow** from the **Capture Mode** list.
            4.  Select **DVB-T** from **DVB Type** under the **Card Selection** group.
            5.  Select **Transponder/Multiplex frequency** to set the transponder frequency.
            6.  Select **Bandwidth** to set the terrestrial bandwidth.
            7.  Select an extra media if you want some background music using **Show more options**. Refer to Playing more than one media file.
            8.  Click on the **Play** button to play the media.
            9.  Click on the **Cancel** button to exit the screen.

        7.  1.  To capture the desktop, select **Open Capture Device** from the **Media** menu. The Open dialog box is displayed.
            2.  Select the **Open Capture Device** tab.
            3.  Select **Desktop** from the **Capture Mode** list.
            4.  Enter a frame rate in the **Desired frame rate** for the capture box.
            5.  Select **Convert** to select the encoding formats and click on the **Save** button. Refer to Converting and Saving a media file format.
            6.  Select an extra media if you want some background music using **Show more options**. Refer to Playing more than one media file.
            7.  Click on the **Play** button to play the media.
            8.  Click on the **Cancel** button to exit the screen.

### Basic Use 0.9 / Playback {#play-howto-basic-use-0-9-playback}

VLC media player helps you to create media files. After creating media files, the quality has to be tested. You can test the quality and several other parameters using playback. In playback, you can specify parameters such as time, bookmarks, and titles.

#### Bookmarks

You can mark and locate particular places in an audio or video file using the Bookmarks feature of VLC. If you want to view a particular scene in a movie or listen to certain tune in a song repeatedly, you can create bookmarks.

To bookmark a scene in a video:

1.  From the *Playback* menu select the *Bookmarks* option, and the *Manage Bookmarks*. The *Edit Bookmarks* dialog box opens.
2.  Click *Create* to create a bookmark for the current track. The created bookmark appears in the *Edit Bookmarks* dialog box.
3.  To view a scene that is bookmarked, select a bookmark from *Bookmarks* in the *Playback* menu.

Edit Bookmarks dialog box under Windows in VLC 1.1.5

#### Title

In a DVD format, each movie is referred to by its title or name. A title is displayed whenever a movie is played by any media player. You can view all titles in a folder in a sequential manner.

1.  To open a folder, select *Open Folder* from the *Media* menu. Locate the folder in which the video files are present and click *OK*.
2.  To select a title, click *Title* in the *Playback* menu. The selected title is then played.

#### Chapter

A video can also be divided into chapters. Different chapters can be accessed at random in a video which is being played. Using this option, you can directly view your favourite chapter without having to see the complete video.

To play a chapter:

1.  Select *Open Folder* from the *Media* menu.
2.  Locate the folder in which the video files are present.
3.  Select a video file and click *OK*.
    The file is played in the VLC media player.
4.  Select *Chapter* in the *Playback* menu to view the list of chapters. Select a chapter of your choice.

Then selected chapter is played.

#### Navigation

In VLC, you can navigate to different titles and their corresponding chapters. You can also customise a DVD by selecting options such as subtitle, angle and so on.

1.  To customize a title, select the required option from *DVD Menu* in the *Navigation* menu.
2.  To view a title, select a *Title* under *Navigation* in the *Playback* menu. The selected title is played.
3.  To view a chapter in a title, select *Title*. When you select a title, the chapters in a title are listed. Select a chapter.

Refer to [#Title and #Chapter sections for more details.

#### Program

This option is enabled only if streams of format DVB and TS are played. Choose the program to select by giving its Service ID. Only use this option if you want to read a multi-program stream (like DVB streams for example). *FIXME: Description needs to be improved*

#### Specify the time

This option is used to go to a specific frame in a media file and listen or view once again.

1.  To specify time select *Jump to Specific Time* from the *Playback* menu. The *Go to Time* dialog box is displayed.
2.  Enter the time in *hh:mm:ss*.
3.  Click on the *Go* button. The control moves the tracker to a specific frame and the media file continues from that specified frame.
4.  Click *Cancel* to exit the dialog box.

Note: Ensure that time limit is within the range of length of the media file.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use 0.9 / Playlist {#play-howto-basic-use-0-9-playlist}

A playlist is a customised list of media files you might want to watch or listen to. Using a playlist, you can specify the media files you want to listen each time you start the VLC media player. You can add tracks from CDs, radio stations, and movies to a playlist. To access the playlist, click on the *Playlist* button in the main interface.

The default playlist view.

#### Additional Sources

In addition to audio and video files, you can play other formats. The additional formats supported by VLC media player are described in the following sections:

-   **Podcasts** - Podcast (Personal On Demand broadCASTING) is a series of audio or video digital media files which is distributed over the Internet and downloaded to media players. If consumers subscribe to Podcasts, whenever new content is added the content gets automatically added to the playlist. You can customise Podcasts. To add a Podcast URL

1.  Select the  *Playlist* button.
2.  Click on the *Internet* button to select it in the left pane. The *Podcasts* menu item will appear under *Internet*.
3.  Select a Podcast stream in the main dialog box. Then right-click the stream and select *Play* from the popup menu.

-   **SAP Announcements** – Helps to advertise your stream over the network.

To play a SAP announcement:

1.  Select the  *Playlist* button.
2.  Click on the *Local Area Network* to select it in the left pane. The *Network Streams (SAP)* menu item will appear under *Local Area Network*.
3.  Select an SAP announcement and right-click. Select *Play* from the popup menu.

-   **Shoutcast Radio Listings** – Shoutcast is a server for streaming the media developed by Nullsoft. Digital audio content can be broadcast from and to media players, and this helps individuals to create Internet radio networks. Using VLC media player, you can listen to your favourite radio stations and you can also create bookmarks to listen to these radio stations in future.

To customize a Shoutcast radio listing:

1.  Select the  *Playlist* button.
2.  Select *Icecast Directory* under *Internet* in the *Playlist* menu. A list of radio stations appears in the right hand panel. If nothing appears in the right hand panel try double-clicking the *Shoutcast Radio* option and wait. It may take a few minutes the first time. After a while, the right hand panel displays a list of titles.

1.  Scroll down and select a radio station.
2.  Right-click on a radio station and:
    1.  Select *Play* if you want to listen to the radio station.
    2.  Select *Remove Selected* if you want to delete the radio station.
    3.  Select the *Stream* option. The Stream output dialog box is displayed. Refer to the Specifying Specifying Streaming options section for more details. Modify the required parameters and click on the *Stream* button to stream the media file.
    4.  Click to select a title in the Playlist dialog box and right-click. Select *Save* from the popup menu. The Stream Output dialog box is displayed. Select the required options and click on the *Save* button in the Stream Output dialog box. Refer to the Specifying Streaming options section for more details.
    5.  Select the Information option. The Media Information dialog box is displayed with details of the media being played.
    6.  Select *Title* to alphabetically sort the radio stations.
    7.  Click to select a title in the Playlist dialog box and right-click. Select *Open Folder* from the popup menu. A folder is opened to show all sub nodes within a title.
    8.  Select *Add Node* to add a node.
    9.  Click to select a title in the Playlist dialog box and right-click. Select *Information* from the popup menu to view the details of the selected title. Refer to the *Media Information* section for more details on options.

-   *Shoutcast TV stream* – You can watch streaming TV using the VLC media player. Shoutcast TV stream refers to a stream transmitted by Nullsoft. The procedure of customising the TV stream and the options are similar to that of the Shoutcast Radio.
-   *Freebox TV listing* – Refers to television service over ADSL accessible by Freebox Free Zone unbundled.

*Note:* You should be connected to the Internet to access these streams.

#### Add Media Files to Playlist

You can add several media files to a playlist. The media files could be selected from the media library, additional sources, and some other source.

To add files to a playlist:

1.  Select the  *Playlist* button.
2.  Right-click on the dialog box and click and a short list appears with two options: Add file and Add directory.
    1.  Select *Add file* to add a file to the playlist.
    2.  Select *Add directory* to add a directory containing media files to the playlist.
3.  Click on the  **Random** icon. This icon toggles between *Random* and *Random Off*. Click on  to play files at random. Click on  and the files are played in an order.
4.  Click on the  *Repeat* icon. This icon toggles between *Repeat One* and *Repeat All*. If you want to listen to a track several times, click on  icon. If you want to listen to all tracks, click on  again.
5.  To search for a media file, enter the name in the *Search* box. To search for media files with certain names or formats, enter a word or phrase in the *Search* box. All files with the specified name are listed.
6.  Click on the  icon. This icon is used to skip to the current item when you have a very long list.
7.  Click the *Remove Selected* button to clear a track from the playlist.

#### Load Playlist

This option is used to add a playlist created in some other media player. You can load playlists of the *.xspf, .asx, .b4s* and *.m3u* formats. To load a playlist:

1.  Select the *Open* option from the *Media* menu. The *Open file* dialog box is displayed.
2.  In the bottom right, change the format to *Playlist Files* in the selector.
3.  Locate a playlist file and click on *Open*.

The selected playlist is added in the current playlist dialog box.

#### Save Playlist

You can save playlists using the VLC media player in format of your choice. To save a playlist:

1.  Create a playlist. Refer to Add Media Files to Playlist for creating a playlist.
2.  Select *Save Playlist to File* from the *Media* menu. The *Choose a filename to save playlist* dialog box is displayed.
3.  Select a name for the playlist.
4.  Select a format in which the playlist must be saved from the *Files of type* list. The Files of type list contains the *.xspf* and *.m3u* formats.
5.  Click on *Save* to save the playlist in the selected format.

#### Play a file

To play a file, open the Media menu, and select the Open File menu item. An Open File dialog box will appear. Select the file you want to open, and click Open. VLC will start playing the selected file. An alternative is to drag 'and' drop your file onto the VLC main interface or playlist window from the file explorer (Finder on MacOS X).

VLC 0.9.8a version Windows XP mode

The File menu - MacOS X interface**- needs verifying for 0.9**

The Open file dialog - wxWidgets interface

(FIXME need 0.9 screenshot for MacOS) The Open file dialog - MacOS X interface

#### Naming Files

You can change the original file name to one you would like before adding the file to VLC. When adding files from the menu bar, the new file name will show in the playlist. However, when dropping the file using the "add/drop" option, VLC may not recognize the name change depending on the file type. In that case, you can right click the header of the playlist column and select "URL," you will then see the original file path for the file.

#### Sorting

In the wxWidgets interface, *Sort* allows you to sort the playlist according to several criteria, or to shuffle it. You can also sort by clicking the header of the column.

In the MacOS X interface, sorting can be done by clicking the header of the column matching the criteria you want to use for sorting.

#### Playlist modes

The playlist supports several playback modes.

In the wxWidgets interface, the toolbar contains three playlist mode buttons. They allow you to enable random mode, to repeat the whole playlist or to repeat one item.

In the MacOS X interface, random mode can be enabled by selecting the *Random* box. A drop down menu allows you to enable playlist and item repeat modes.

#### Misc

##### Search

You also have a search tool. Enter a search string and hit search. The next item to match the string will be highlighted. Keep hitting Search to cycle between all matching items.

##### Moving items

In the wxWidgets interface, the *Up* and *Down* buttons at the bottom of the playlist window allow you to move an item. Select an item and use these buttons to move it.

In the MacOS X interface, you can easily move an item with the mouse, using drag-and-drop.

##### Contextual menu

By right-clicking or control-clicking an item, a contextual menu will appear, giving access to a number of functions (for example, play the item, disable it, delete it, or get info on it).

##### Example finding a Shoutcast radio stream

This example was verified as working on 15 October 2008, using VLC 0.9.4 under Windows Vista. *This needs reproducing by other people on other versions and other operating systems.*

1\. Ensure your firewall is set to allow the VideoLan program to make outgoing connections.

2\. Click *Tools* then *Preferences*, click Interface and then click All under "Show settings". Then click the "-" next to "Playlist" in order to show the "Services discovery" submenu. If the shoutcast radio listings box is empty, click it so that a check-mark appears. The text field underneath should now show the word "shout". Click the Save button to save and close the Preferences window:

3\. Restart VLC media player to make it take notice of the changed preferences.

4\. On the VLC interface click *Playlist*, then click *Show Playlist*. Select the "Shoutcast Radio" in the left hand panel. If nothing appears in the righthand panel, try double-clicking "Shoutcast Radio" and waiting, it may take a few minutes the first time. After a while the righthand panel displays a long list of titles.

5\. Scroll down the radio stations in the right-hand panel and select one. Click the mouse right button and click the "Play" item.

6\. It may take some time for the connection to the radio station to establish (and it may fail if the station's outgoing streams are all occupied). When it does connect, VLC should start playing the audio stream from the station:

##### Example playing a known Shoutcast radio stream

Go to 0 and search for a radio station of your choice. On Windows, right-click your mouse over Shoutcast's "Tunein" button and click "Save Link As..." to save the playlist on your computer. Remember where you saved the playlist, rename it to something that makes sense.

At any time later, you can use VLC to open the saved playlist and listen to that radio station.

For example, to find a BBC World Service radio stream, use a browser to go to: 0

One of the stations listed may be playing the World Service, if so move your mouse over the "TUNEIN!" webicon and click the right mouse button and click "Save Link As...", as described above.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use 0.9 / Snapshots {#play-howto-basic-use-0-9-snapshots}

There are two ways to take snapshots (i.e., screenshots or frame grabs) with VLC:

1.  Open the *Video* menu, and select the *Take Snapshot* menu item.
2.  Press the snapshot hotkey
    -   Linux / Unix / Windows (Qt Interface): Shift+s
    -   Mac OS X: Command+Alt+s

When a snapshot is taken, it will briefly preview as a thumbnail with its filename and then fade away.

To change the hotkey, go to Tools → Preferences. If "Show settings" is set to Simple, click Hotkeys; if "Show settings" is set to All, navigate to Interface → Hotkeys settings. Set the hotkey for Take video snapshot.

#### Snapshot location, format and name

The snapshot location depends upon your operating system:

-   Windows XP: "%HOMEPATH%\\My Pictures\\"
-   Windows Vista, 7, 8, and 10: "%HOMEPATH%\\Pictures\\"
-   Linux / Unix: \~/Pictures
-   macOS: Desktop/

##### Configuring snapshot options under Windows:

The location, format and name of snapshots may be changed in the *Preferences* menu item in the *Tools* tab, subsection *Video*.

The default format for snapshots is PNG, but this may be changed to JPEG. Also, the default name for snapshots is *vlcsnap-* followed by a timestamp that is *not* the time of the frame in the video you're viewing, but rather the current date and time—as in 2014-01-16-14h57m19s163.

Also, you may substitute other text for *vlcsnap-* in the *Video snapshot file prefix* and you may choose to have snapshots numbered sequentially (i.e., 000001, 000002, 000003, and so on) instead of with a timestamp.

As of version 0.9.0, you may even use variables in the text used for the filename. For example, *\$T* (must be upper case) will insert the video's time code into the file name. If you were to change the prefix to *Friends-\$T-* while watching a DVD of *Friends*, then the snapshot filenames would look something like this: *Friends-00_05_21-2014-01-16-14h57m19s163.png*. This indicates a snapshot taken at 5 minutes and 21 seconds into the video; and it was taken on this day at this time: *2014-01-16-14h57m19s163*.

For a shorter file name, check the "Sequential numbering" option in the configuration box (below). Instead of numbers like *2014-01-16-14h57m19s163*, VLC will simply insert the count of snapshots for that session—for example, *00004*. Thus, in the example above, a snapshot with sequential numbering would look like this: *Friends-00_05_21-000001.png*

For a full list of variables, please see Documentation:Play HowTo/Format String.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use 0.9 / Subtitles {#play-howto-basic-use-0-9-subtitles}

VLC supports many kinds of subtitles.

#### Media with included subtitles

Many types of media can have embedded subtitles. VLC can read subtitles for the following media formats:

-   DVD
-   SVCD
-   OGM files
-   Matroska (MKV) files

Subtitles are enabled by default in VLC media player. To disable them, go to the *Video* menu, and to *Subtitles track*. All available subtitles tracks will be listed. Select "Disable" to turn off the subtitles. Depending on the media, a description (language, for example) might be available for the track.

To disable subtitles by default, select "Preferences", then "Show All". Select "Input/Codecs". On the "Subtitle Track ID" selection window, change the value to "-1". (NOTE: Changing the value in the "Subtitle Track" menu will not disable the subtitle file.) In the case of multiple subtitle tracks, a value of "0" will enable subtitle track 1, a value of "1" will enable subtitle track 2, and so on.

VLC under Linux:

VLC under OSX:

VLC under Windows:

DVD and SVCD subtitles are merely images, so you won't be able to change anything for them. OGM and Matroska subtitles are rendered text, so you will be able to change several options.

Text rendering options can be changed in the *Preferences* in the *Tools* tab. To adjust the font preference check the *All* bullet in the *Show Settings* box, and then click *Subtitles/OSD*. You can then set the font and its size under *Text Renderer*. For the font, you have to select a font file. In Windows, they can be found in *C:\\Windows\\Fonts*. Under MacOS X, they are in */System/Library/Fonts*. Sizes can be set either relatively or as a number of pixels.

Subtitle text rendering preferences under Windows, VLC 1.1.5

You need to restart your stream for the font modifications to take effect.

#### Subtitles files

While modern file formats like Matroska or OGM can handle subtitles directly, older formats like AVI can't. Therefore, a number of subtitles files formats have been created. You need two files: the video file and the subtitles file that only contains the text of the subtitles and timestamps.

VLC can handle these types of subtitles files:

-   MicroDVD
-   SubRIP
-   SubViewer
-   SSA
-   Sami
-   Vobsub (this one is quite special: it is not made from text but from images, which means that you can't change the fonts)

To open a subtitles file, use the Advanced Open dialog box (Menu File, Open file). Select your file by clicking on the *Browse* button. Then, check the *Subtitle options* checkbox and click on the "Settings" button.

You can then select the subtitles file by clicking the *Browse* button. You can also set a few options like character encoding, alignment and size.

An alternative is loading subtitles from the *Subtitles Track* menu item under the *Video* tab.

Note: For Vobsub subtitles, you need to select the **.idx** file, not the **.sub** file. Encoding, alignment and size won't have any effect for Vobsub subtitles.

Font can be changed as explained in the previous section.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use 0.9 / Video {#play-howto-basic-use-0-9-video}

You can play video files, video clips and other video media using the VLC media player. You can resize, change aspect ratio, crop videos, load subtitles, deinterlace, save snapshots, and convert videos to DirectX wallpaper.

Video tracks of the *.asf, .avi, .divx, .dv, .mxf, .ogg, .gm, .ps, .ts, .vob,* and *.wmv* formats are supported.

#### Playing a Video Track

There are two main ways to open and play a video track:

1.  Select *Open File* from the *Media* menu.

     2. Select a video track and double-click it or click the *Open* button.

The selected track will play.

#### Loading Subtitle Tracks

A subtitle is a textual version of a movie’s dialogue. Subtitles are helpful if you are viewing a movie that contains foreign language(s). You can load subtitles for video tracks. Subtitles of the formats *.cdg, .idx, .srt, .sub, .utf, .ass, .ssa, .aqt, .jss, .psb, .rt* and *smi* are supported.

VLC can read subtitles for the media formats such as *DVD*, *SVCD*, *OGM* files, and *Matroska (MKV)* files.

To enable the subtitle for a track:

1.  Select *Open File* under the *Subtitle* menu item from the *Video* menu. The *Open Subtitles File* dialog box is displayed.  
2.  Locate the file which contains the subtitle and click on *Open*. The subtitles are displayed.

For more details, see [Documentation:Subtitles](#subtitles).

#### Full Screen

This option is useful if you want to watch the video in the full screen mode.

1.  Select *Full Screen* from the *Video* menu. The video will then occupy the entire screen.
2.  To return to the original mode, press *Esc* on the keyboard or right-click the mouse and select the *Leave Full Screen* option. The video will then return to its original mode.

Note: When you switch to full screen, the controls may appear for a short period of time. To restore the controls after they disappear, move the mouse or press any key on the keyboard.

#### Always on Top

This option is useful if you want the VLC media player to remain on the top of the screen always when other applications or files are open.

1.  To make the VLC media player appear on top of the screen, select *Always on Top* from the *Video* menu. 
2.  If you do not want VLC to appear on the top of the screen, select the *Always on Top* option from the *Video* menu and manually minimise the VLC application.

#### DirectX Wallpaper

This option is useful if you want to display the video which is being played as your desktop wallpaper.

To view the current video file as wallpaper

1.  Select *Advanced File Open* from the *Media* menu. The *Open Media* dialog box is displayed. 
2.  Select a file and click  *Play*.
3.  Select *DirectX Wallpaper* from the *Video* menu.

The wallpaper mode will then display the video as the desktop background.

Note: that this feature works only if you deactivate the overlay under Windows XP.

#### Snapshot

This option is useful if you want to capture a portion of the video as an image.

1.  Select *Advanced File Open* from the *Media* menu. The Open dialog box is displayed.
2.  Select a file and click  *Play*.
3.  To capture an image from the video, select *Snapshot* from the *Video* menu.

The image is captured in the *.png* picture format and is saved in the *C:\\My Pictures* folder by default (*C:\\Users\\**Username**\\Pictures*).

#### Zoom

You can enlarge videos in different sizes. This option is useful if you want to change the size of a video track which is being played. The supported sizes are *1:4 Quarter, 1:2 Half, 1:1 Original (default)* and *2:1 Double*.

To view a video in a particular dimension, select a dimension from *Zoom* in the *Video* menu. The track is then resized based on the selected zoom ratio.

#### Aspect Ratio

Aspect ratio refers to the width of a picture in relation to its height. For example, the ratio 4:3 means four units wide to three units high. VLC provides a list of aspect ratio values which are *Default, 1:1, 4:3, 16:9, 16:10, 2.21:1, 2.35:1, 2.39:1* and *5:4*.

To select an aspect ratio, select a value from *Aspect Ratio* in the *Video* menu. The video is then adjusted based on the selected ratio.

#### Crop

This option is helpful if you want to capture a small portion of a video as an image. This also helps crop the black bars of the top and bottom of a video.

The cropping values that are supported are *Default, 16:10, 16:9, 1.85:1, 2.21:1, 2.35:1, 2.39:1, 5:3, 4:3, 5:4,* and *1:1*.

To crop a video that is played, select a value from *Crop* in the *Video* menu. The video is then cropped based on the selected value.

#### Deinterlace

Deinterlace refers to a process where interlaced video signals are converted into non-interlaced signals. VLC provides the *Discard, Blend, Mean, Bob, Linear, X, Yadif and Yadif (2x)* deinterlacement modes.

1.  Select *Deinterlace* from the *Video* menu and choose the appropriate setting.
2.  To change the deinterlacement mode select 'Deinterlace mode' is the *Video Menu*
3.  Select a mode and observe the change in the video being played.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use 0.9 / Video and Audio Filters {#play-howto-basic-use-0-9-video-and-audio-filters}

This page is outdated and information might be incorrect.

VLC includes a system of *filters* that allow you to modify the audio and video.

#### Deinterlacement and Post Processing

VLC is able to deinterlace a video stream using different deinterlacement methods. Deinterlacement can be enabled in the *Video* menu, *Deinterlacement* menu item. The *Blend* methods gives the best results in most cases. The *discard* method is a less resource consuming alternative, although its results may be slightly compromised.

On some particular streams (MPEG 4, DivX, Xvid, Sorenson, etc.), some additional image filtering can be applied to the video before display, improving its quality in some cases. This can be enabled by using the *Post processing* menu item in *Video*. Different levels of post processing can be chosen here. A higher level means more filtering.

#### Video filters

##### Summary

VLC features several filters able to change the video (distortion, brightness adjustment, motion blurring, etc.).

In Windows and Linux, the user must go to the *Effects and Filters* in the *Tools* menu item. A dialogue box entitled "Adjustments and Effects" will appear.

In macOS you can enable these filters through the *Extended Controls panel*. Click on the triangle next to *Video filters* to select your filters or expand the *Adjust Image* section to change the contrast, hue, etc.

iOS:

Example of combined effects on a video:

##### Rotate

You can easily rotate a video. Open the *Effects and Filters* dialog, in the *Tools menu*

Select the *Video Effects* tab, then the *Geometry* one.

Check the *Transform* checkbox to use rotation presets (90°, 180°, 270°) or check the *Rotate* checkbox to manually select the angle you wish to apply.

#### Audio filters

##### Equalizer

[Wikipedia](http://en.wikipedia.org/wiki/Main_Page) has information on this entry:

***[Equalization (audio)](http://en.wikipedia.org/wiki/Equalization_(audio) "wikipedia:Equalization (audio)")***

VLC features a 10-band graphical equalizer, a device used to alter the relative frequencies of audio (e.g. for a bass boost). You can display it by activating the advanced GUI on wxWidgets or by clicking the *Equalizer* button on the macOS interface. The following image is the interface of the audio equalizer in the Windows and GNU/Linux interface.

The equalizer in the macOS interface

Presets are available in all of these dialog boxes.

##### Other audio filters

At the moment, VLC features two other audio filters: a volume normalizer and a filter providing sound spatialization with a headphone. They can be enabled in the *Effects and Filters* menu item in the *Tools* tab of the Windows and GNU/Linux interface and in the Audio section of the Extended Controls panel of the macOS interface.

For better control, you need to go to the preferences. To select the filters to be enabled, go to *Audio*, then to *Filters*. In the "audio filters" box, enter the names of the filters to enable, separated by commas. Valid names are "equalizer", "normvol" and "headphone".

If you want to tune the behavior of these filters, go to *Audio, Filters, \[your filter\]*. The equalizer and headphone filters can be tuned.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use / Audio {#play-howto-basic-use-audio}

VLC can play several audio formats: *.asf, .avi, .divx, .dv, .mxf, .ogg, .gm, .ps, .ts, .vob,* and *.wmv*. It can convert audio tracks and use several visualizations.

**Note:** The commands in the **Audio** menu are only enabled when an audio file is being played.

#### Playing an audio track

To play a track:

1.  Select *Open File* in the *Media* menu.
2.  Select an audio file and click on the  *Play* button. The selected track is played.

#### Enabling and disabling audio tracks

-   To disable a track, select the *Disable* option in the *Audio Track* from the *Audio* menu. The selected track will then stop.
-   To enable the track again, select the designated *Track* option in the *Audio Track* from the *Audio* menu. The selected track will then play.

#### Recording Audio

To record audio you need the record button () to be visible. The record button is hidden by default. You can display using one of these methods:

-   Select Advanced Controls in the View menu. The Advanced toolbar is displayed on top of the standard toolbar. The Advanced toolbar contains the Record button.
-   Select Customize interface in the Tools menu and add the record button to the Line 2 of buttons (which is the line shown by default).

Once the Record button is visible, click it to start recording.

The recording from a shoutcast stream is stored somewhere in your files under a name like 0 (e.g.: 1, when recording from [Radio CAFF](http://radiocaff.com.ar/) (or more precisely from the underlying [WinAmp stream](http://panel7.serverhostingcenter.com/tunein.php/radiocaff/playlist.pls)). Under my german Windows XP it was stored under "Eigene Dateien/Eigene Music" so I guess that you find it in an english Windows under "My Documents/My Music/", I don't know where it will be stored under Linux or any other OS (updates are welcome).

You can automagically cut the stream into tracks by relaying the stream through [Streamripper](http://streamripper.sourceforge.net), i.e. by directing StreamRipper to the ShoutCast stream and directing VLC to the relaying port of StreamRipper (default http://localhost:8000).

#### Audio Device

This option helps you to listen to audio files in two modes: stereo and mono.

1.  To listen to an audio track in either the Stereo or Mono mode, select *Open File or Open Disc* from the *Media* menu. The Open dialog box is displayed.
2.  Select an audio file and click on the  *Play* button. The selected track is played.
3.  Select *Mono* in *Audio Device* from the *Audio* menu if you want to listen to the audio track in the Mono mode.

Mono refers to monaural sound that uses a single channel for sound reproduction.

1.  Select *Stereo* in *Audio Device* from the *Audio* menu if you want to listen to the audio track in the Stereo mode.

Stereo refers to sound that uses two channels for sound reproduction or stereophonic sound.

#### Audio Channels

In audio, a channel refers to a stream of audio that is to be played by one speaker. For example, stereo audio, consists of two channels. This option is useful for codecs that don’t have support for more than 2 channels.

Select a channel type in *Audio Channels* from the *Audio* menu. VLC media player provides four audio channels and they are:

1.  *Stereo* – Refers to the reproduction of the sound in two or more independent audio channels using more than one speaker. If you use this option, you would feel as though the sound is played from all the directions. You can observe this in a regular home theatre with 5.1 or 6.1 speakers.
2.  *Left* – You can observe this in a regular audio player with 2.1 speakers. If you select the **Left** option, the music is played only in the left speaker. The speaker on your right is automatically switched OFF.
3.  *Right* - If you select the **Right** option, the music is played only in the speaker on your right side. The speaker on your left is automatically switched OFF.
4.  *Reverse Stereo* – There are several applications that are used to reverse the stereo whereas VLC has an in-built feature to reverse the stereo. This option is useful if you want the audio to play in tandem with the video. You can use the **Reverse Stereo** option if you want to deliberately change the audio output.

Imagine that you are watching a video. In the video, a person walks on the left side but the sound is produced on the right speaker. You can correct this by selecting the *Reverse Stereo* option in VLC. Select the *Reverse Stereo* option and play the same scene in the video and observe the difference.

You can observe this with 2.1, 5.1, 6.1 and 8.1 speakers.

#### Visualize Audio

Visualizations display splashes of colour and geometric shapes and generate animated imagery based on a piece of music.

The different visual effects available are *Spectrometer, Scope, Spectrum, VU Meter and Goom*. This menu item can also be used to disable a visualization.

1.  Select an option under the *Visualizations* option from the *Audio* menu to view the effects. The selected visualization is then played.
2.  To disable visualizations, select *Disable* under *Visualizations* from the *Audio* menu. The visualization is then disabled.

Spectrum visualization on VLC:

#### Maximum VLC Volume

To change the maximum volume in % that VLC should use, go to **Tools** → **Preferences** (select **All** at bottom left corner) → **Interface** → **Main interfaces** → **Qt** → **Maximum volume displayed**.

Save it and restart VLC.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use / Basic troubleshooting {#play-howto-basic-use-basic-troubleshooting}

**Languages:English** • français

#### VLC Support Guide: Solve your VLC issues right now!

The **V**LC **S**upport **G**uide is an informal, step-by-step guide for troubleshooting most common issues with VLC.

It complements the VLC media player Documentation.

**So what's your problem?**

##### Installation Issue

VLC won't install. Go!

##### Startup Issue

VLC won't start up. Go!

##### Audio Playback Issue

The audio or the sounds are wrong. Go!

##### Video Playback Issue

The video is messed up. Go!

##### Subtitle Display Issue

The subtitles aren't working properly. Go!

##### Usage Issue

I have difficulty using VLC. Go!

##### Interface Issue

I want to change my interface. Go!

##### Uninstallation Issue

VLC won't uninstall (why are you uninstalling it anyway?). Go!

#### Get Help

If this troubleshooter does not resolve your problems or answer your questions, some other resources which you can use include:

-   Frequently asked questions
-   Frequently asked questions about VLC on Windows
-   Frequently asked questions about VLC on macOS
-   Frequently asked questions about VLC on Linux
-   The [VideoLAN support forum](https://forum.videolan.org/)
-   The VideoLAN IRC channel.
-   VLC documentation

This page is part of the informal VLC Support Guide.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use / Hotkeys {#play-howto-basic-use-hotkeys}

Most of VLC functions are accessible using hotkeys.

The list of the available hotkeys and their functions can be retrieved and altered in the *Preferences* panel of the player. In the Windows and Linux interface, *Preferences* are available in the "Tools" tab as the "Preferences" menu item. In the MacOS X interface, open the "VLC" menu, and select "Preferences". Select the "Hot keys" panel in the dialog.

As of version 0.9, a list of hotkeys are presented in a drop-down window. To change one, double-click its name to select it. Then, press the new key that will trigger the specified action. Modifier keys (such as Control/Command and Alt) may also be used. In the 1.x version you can also filter hotkeys with a search filter.

In earlier versions, several boxes gave the list of modifiers for the hotkey. To trigger an action using a hotkey, you need to press simultaneously the keys corresponding to the different selected modifiers as well as the key set in the dropdown.

To change the binding of a hotkey, select or deselect boxes corresponding to the different modifiers, and change the key by using the drop-down menu. Select the *Save* button to apply the changes.

The Hotkeys Panel - MacOS X interface**FIXME - needs verifying for 0.9**

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use / Interface {#play-howto-basic-use-interface}

#### General Interface Description

VLC has several interfaces:

-   A cross-platform interface for Windows and GNU/Linux, which is called Qt.
-   A native Mac OS X interface.
-   An interface that supports skins for both Windows and GNU/Linux.

The operation of VLC is essentially the same in all the interfaces.

##### Windows and GNU/Linux (Qt)

The screenshot below shows the default interface in VLC 2.0. More features can be displayed by selecting them in the *View* menu.

See also VLC Interface 2.0 on Windows 7

##### Mac OS X

This screenshot shows the default interface that VLC had on Mac OS X until version 1.1:

Since version 2.0 the interface has been redesigned. See OSX 2.0 interface.

#### Starting VLC Media Player in Windows

In Windows XP: Click **Start** -\> **Programs** -\> **VideoLAN** -\> **VLC media player**.

In Windows 7: Click **Start** -\> **All Programs** -\> **VideoLAN** -\> **VLC media player**.

VLC is shown on the screen and a small icon  is shown in the system tray.

#### Stopping VLC Media Player

There are three ways to quit VLC:

-   Right click the VLC icon () in the tray and select **Quit** (*Alt-F4*).
-   Click the **Close** button in the main interface of the application.
-   In the **Media** menu, select **Quit** (*Ctrl-Q*).

#### Notification Area Icon

Clicking this icon shows or hides the VLC interface. Hiding VLC does not exit the application. VLC keeps running in the background when it is hidden. Right clicking the icon in the notification area shows a menu with basic operations, such as opening, playing, stopping, or changing a media file.

#### Main Interface

The main interface has the following areas:

-   **Menu bar**.
-   **Track slider** - The track slider is below the menu bar. It shows the playing progress of the media file. You can drag the track slider left to rewind or right to forward the track being played. When a video file is played, the video is shown between the menu bar and the track slider.
    **Note: When a media file is streamed, the track slider does not move because VLC cannot know the total duration.**
-   **Control Buttons** - The buttons below the track slider cover all the basic playback features.

Click here to view an explanation of every menu item.

#### Opening media

See Documentation:Play HowTo/Basic Use 0.9/Opening modes

#### Streaming Media Files

Streaming is a method of delivering audio or video content across a network without the need to download the media file before it is played. You can view or listen to the content as it arrives. It has the advantage that you don't need to wait for large media files to finish downloading before playing them.

VideoLan is designed to stream MPEG videos on high bandwidth networks. VLC can be used as a server to stream MPEG-1, MPEG-2 and MPEG-4 files, DVDs and live videos on the network in unicast or multicast. Unicast is a process where media files are sent to a single system through the network. Multicast is a process where media files are sent to multiple systems through the network.

VLC is also used as a client to receive, decode and display MPEG streams. MPEG-1, MPEG-2 and MPEG-4 streams received from the network or an external device can be sent to one machine or a group of machines.

**To stream a file**:

1.  From the **Media** menu, select **Open Network Stream**. The *Open Media* dialog box loads with the *Network* tab selected.
2.  In the **Please enter a network URL** text box, Type the network URL.
3.  Click **Play**.

Note: When VLC plays a stream, the track slider shows the progress of the playback.

For more information, refer to Documentation:Streaming HowTo/Receive and Save a Stream

#### Converting and Saving a Media File Format

VLC can convert media files from one format to another.

**To convert a media file**:

1.  From the **Media** menu, select **Convert/Save**. The *Open media* dialog window appears.
2.  Click **Add...**. A file selection dialog window appears.
3.  Select the file you want to convert and click **Open**. The *Convert* dialog window appears.
4.  In the **Destination file** text box, indicate the path and file name where you want to store the converted file.
5.  From the **Profile** drop-down, select a conversion profile.
6.  Click **Start**.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use / Interface in Windows 7 {#play-howto-basic-use-interface-in-windows-7}

-

### Basic Use / Interface OSX {#play-howto-basic-use-interface-osx}

This page is outdated and information might be incorrect.

** VLC 1.2 Mac Interface Documentation**

Below image shows all the functions of the buttons present in VLC 1.2 Mac Version

"Fullscreen mode" for Snow Leopard and Leopard is located on the bottom of the window.

Below shows the Fullscreen interface and functions of each button used during the Fullscreen mode.

** Basic Playlist Controls in VLC 1.2 for Mac**

This is an image of the first view one will get on opening VLC (the submenus to the left, are by-default closed).

Media can be added to the playlist by clicking on the Open media... button, and choosing your options appropriately.

Alternately, media can be added by dragging and dropping its icon from anywhere into the box.

Additional media can be added in between the playlist, in any desirable order, or into a new playlist, by dragging and dropping further.

### Basic Use / Interface Windows {#play-howto-basic-use-interface-windows}

**VLC 1.2 WIndows Interface Documentation
**

**Below image shows all the functions of the buttons present in VLC 2.0 Windows Version**

**Below shows the Fullscreen interface and functions of each button used during the Fullscreen mode.**

**Various visualization in VLC 1.2 Windows**

projectM has been removed as of VLC 2.2.0. See [this thread](https://forum.videolan.org/viewtopic.php?f=14&t=124958&p=425222&hilit=projectM#p425222).

**This is an image of the first view one will get on opening VLC (the submenus to the left, are by-default closed).**

**Advanced way of Opening,Saving,Converting,Streaming files**

**VLC skinned version, can be accessed from the start menu**

**VLC skinned version, fullscreen interface**

### Basic Use / Interface Windows 7 {#play-howto-basic-use-interface-windows-7}

VLC 1.2 Windows 7 Interface Documentation

### Basic Use / Menus {#play-howto-basic-use-menus}

The following table outlines every option in the menus of VLC Media Player:

-   Menu: **OptionDescriptionHotkey**
-   Media: **Open File** Select this option to open a media file and play it. *Ctrl+O*
-   : **Advanced Open File** Select this option to open files through a folder, disc, network or a capture device. *Ctrl+Shift+O*
-   : **Open Folder** Select this option to open a folder with multiple media files and play them in order. *Ctrl+F*
-   : **Open Disc** Select this option to open discs with different video format. *Ctrl+D*
-   : **Open Network Stream** Select this option to receive media files from Internet and play them. *Ctrl+N*
-   : **Open Capture Device** Select this option to receive media files from capture devices such as camcorder, webcam and so on. *Ctrl+C*
-   : **Open Location from clipboard** Select this option to enter the URL or path to the media you want to play. *Ctrl+V*
-   : **Recent Media** Scroll over this option to view recently played media. *-*
-   : **Save Playlist to File** Select this option to save the playlist to a file. *Ctrl+Y*
-   : **Convert/Save** Select this option to convert media files to different media file formats. *Ctrl+R*
-   : **Streaming** Select this option to send media files through network and play them live. *Ctrl+S*
-   : **Quit** Select this option to quit the application. *Ctrl+Q*
-   Playback: **Bookmarks** Select this option to bookmark the media file. *-*
-    : **Title** Select this option to randomly access a particular movie in a DVD. *-*
-    : **Chapter** Select this option to randomly access a particular chapter in a movie. *-*
-    : **Navigation** Select this option to navigate to different titles and their corresponding chapters. *-*
-    : **Program** To be added *-*
-   : **Jump to Specific Time** Select this option to move the track slider to a specific frame of the video or audio file. The media file starts playing from that instance. *Ctrl+T*
-   Audio: **Audio Track** Select this option to disable or enable an audio track. *-*
-   : **Audio Channel** Select this option to select a audio channel. *-*
-   : **Audio Device** Select this option to convert the stereo audio files to mono and vice-versa. *-*
-    : **Visualization** Select this option to display splashes of colour and geometric shapes while listening to an audio file. *-*
-   Video: **Video Track** Select this option to disable or enable a video track. *-*
-    : **Subtitles Track** Select this option to load subtitle files for video files requiring subtitles. *-*
-    : **Full Screen** Select this option to view the media file in the entire screen. *-*
-    : **Always On Top** Select this option to display the application always on top when other applications or files are open. *-*
-    : **DirectX wall paper** Select this option to make the video which is being displayed to become the wall paper. *-*
-    : **Snapshot** Select this option to capture snapshots of video being displayed. *-*
-    : **Zoom** Select this option to zoom in or zoom out a video track. *-*
-    : **Scale** Select this option to resize the video to an appropriate size. *-*
-    : **Aspect Ratio** Select this option to adjust the width to height ratio of the video. *-*
-    : **Crop** Select this option to crop the edges of a video track. *-*
-    : **Deinterlace** Select this option to convert interlaced video signals into non-interlaced form. *-*
-    : **Deinterlace mode** Select this option to choose the type of deinterlacing. *-*
-    : **Post processing** Select this option to disable post processing or the processing number. *-*
-   Tools: **Effects and Filters** Select this option to adjust the audio and video effects. *Ctrl+E*
-    : **Track Synchronization** Select this option to display a menu of options to have audio and video effects, and synchronization of the audio and video media. *-*
-    : **Media Information** Select this option to view information regarding the media file being played. *Ctrl+I*
-    : **Codec Information** Select this option to view codec information of the media file being played. *Ctrl+J*
-    : **Bookmarks** Select this option to bookmark the media file. *Ctrl+B*
-    : **VLM Configuration** To be added *Ctrl+W*
-    : **Program Guide** Select this option to view the program schedule *-*
-    : **Messages** Select this option to view the messages and the modules tree. *Ctrl+M*
-    : **Plugins and extension** Select this option to view plugins and extension installed on VLC. *Ctrl+M*
-    : **Preferences** Select this option to give preferred settings to audio, video, subtitles, codecs and hotkeys. *Ctrl+P*
-   View: **Playlist** Select this option to view the playlist. *Ctrl+L*
-    : **Add Interface** Select this option to change the interface. *-*
-    : **Minimal View** Select this option to enable minimal view (removes track list and basic control buttons). *Ctrl+H*
-    : **Fullscreen Interface** Select this option to toggle fullscreen mode. *F11*
-    : **Advanced Controls** Select this option to toggle advance controls below the tracklist. *-*
-    : **Docked Playlist** To be added *Ctrl+W*
-    : **Customize Interface** Select this option to adjust the interface. *-*

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use / Open {#play-howto-basic-use-open}

#### Play a file

To play a file, open the *Media* menu, and select the *Open File* menu item.

An *Open File* dialog box will appear. Select the file you want to open and select *Open*.

VLC will then start playing the designated file. An alternative is to simply drag 'n' drop your file into the VLC main interface or the playlist window from the file explorer (Finder on Mac OS X).

#### Play a CD/DVD/VCD

To play a CD, VCD or a DVD, open the *Media* menu and select *Open Disc* menu item. In the *Open Disk* dialog box, select the type of media (DVD, SVCD/VCD or Audio CD).

You can either select the drive in which the media is located by selecting the drive letter from the *Disc Device* drop-down list, or you can select the *Browse* button, which will open a dialog box that you can use to browse for the media you wish to play.

If you want to start the DVD or VCD playback from a given title and chapter instead of from the beginning, you can set it using the *Title* and *Chapter* selectors. You can also set the *Audio* and *Subtitles* track using the selectors. There is also an option for *No DVD menus*, when reading a DVD.

To start playback select the *Ok* button.

#### Play a network stream (WebRadio, WebTV, etc.)

To open a network stream, open the *Media* menu and select the *Open Network Stream* menu item.

A dialog box will then open with three user input boxes. The first one is for the user to select the *Protocol* of the stream that they wish to open (HTTP/HTTPS/MMS/FTP/RTSP/RTP/UDP/RDMP). The second box is for the user to input the *Address* of the stream and the third one is for the user to select the appropriate port. However in the latest version of VLC (1.1.5), the user only needs to input the *Address* (examples are shown in image above).

To begin playback, select the *Play* button.

If you get some stuttering during playback, you can try to increase the size of the read buffer. This can be done in the *Open Network Stream* dialog box, by firstly checking the *Show more options* check box then adjusting the *Caching* selector, which allows you to choose the amount of time (in milliseconds) VLC should store data in its buffer before starting playback.

#### Play from an acquisition card

To play from an acquisition open the *File* menu, and select *Open Capture Device*.

From here you can choose the *Capture Mode* and the *Video/Audio Device Name*. The user can also adjust the configuration for these devices by clicking *Configure*. The user is also able to set the size of the video that will be played by the Direct Show plugin and options such as 'Device Properties' and 'Tuner Properties' by clicking *Advanced Options*.

For Video4Linux devices, you can set the name of the video and audio devices using the "Video device name" and "Audio device name" text inputs. The "Advanced options..." button allows you to select some further settings useful in some rare cases, such as the chroma of the input (the way colors are encoded) and the size of the input buffer.

To use a Hauppauge PVR card, select the PVR tab in the "Open" dialog box. Use the "Device" text input to set the device of the card you want to use. You can set the Norm of the tuner (*PAL, SECAM or NTSC*) by using the "Norm" Drop Down. The Frequency selector allows you to set the frequency of the tuner (in kHz), the bitrate selector to set the bitrate of the resulting encoded stream (in bit/s). The "Advanced Options button allows to set some more settings, such as the size of the encoded video (in pixels), its framerate (in frame per second), the interval between 2 key frames, etc.

To start playback from an acquisition card, click *Play*.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use / Playback {#play-howto-basic-use-playback}

VLC media player helps you to create media files. After creating media files, the quality has to be tested. You can test the quality and several other parameters using playback. In playback, you can specify parameters such as time, bookmarks, and titles.

#### Bookmarks

You can mark and locate particular places in an audio or video file using the Bookmarks feature of VLC. If you want to view a particular scene in a movie or listen to certain tune in a song repeatedly, you can create bookmarks.

To bookmark a scene in a video:

1.  From the *Playback* menu select the *Bookmarks* option, and the *Manage Bookmarks*. The *Edit Bookmarks* dialog box opens.
2.  Click *Create* to create a bookmark for the current track. The created bookmark appears in the *Edit Bookmarks* dialog box.
3.  To view a scene that is bookmarked, select a bookmark from *Bookmarks* in the *Playback* menu.

Edit Bookmarks dialog box under Windows in VLC 1.1.5

#### Title

In a DVD format, each movie is referred to by its title or name. A title is displayed whenever a movie is played by any media player. You can view all titles in a folder in a sequential manner.

1.  To open a folder, select *Open Folder* from the *Media* menu. Locate the folder in which the video files are present and click *OK*.
2.  To select a title, click *Title* in the *Playback* menu. The selected title is then played.

#### Chapter

A video can also be divided into chapters. Different chapters can be accessed at random in a video which is being played. Using this option, you can directly view your favourite chapter without having to see the complete video.

To play a chapter:

1.  Select *Open Folder* from the *Media* menu.
2.  Locate the folder in which the video files are present.
3.  Select a video file and click *OK*.
    The file is played in the VLC media player.
4.  Select *Chapter* in the *Playback* menu to view the list of chapters. Select a chapter of your choice.

Then selected chapter is played.

#### Navigation

In VLC, you can navigate to different titles and their corresponding chapters. You can also customise a DVD by selecting options such as subtitle, angle and so on.

1.  To customize a title, select the required option from *DVD Menu* in the *Navigation* menu.
2.  To view a title, select a *Title* under *Navigation* in the *Playback* menu. The selected title is played.
3.  To view a chapter in a title, select *Title*. When you select a title, the chapters in a title are listed. Select a chapter.

Refer to #Title and #Chapter sections for more details.

#### Program

This option is enabled only if streams of format DVB and TS are played. Choose the program to select by giving its Service ID. Only use this option if you want to read a multi-program stream (like DVB streams for example). *FIXME: Description needs to be improved*

#### Specify the time

This option is used to go to a specific frame in a media file and listen or view once again.

1.  To specify time select *Jump to Specific Time* from the *Playback* menu. The *Go to Time* dialog box is displayed.
2.  Enter the time in *hh:mm:ss*.
3.  Click on the *Go* button. The control moves the tracker to a specific frame and the media file continues from that specified frame.
4.  Click *Cancel* to exit the dialog box.

Note: Ensure that time limit is within the range of length of the media file.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use / Playlist {#play-howto-basic-use-playlist}

A playlist is a customised list of media files you might want to watch or listen to. Using a playlist, you can specify the media files you want to listen each time you start the VLC media player. You can add tracks from CDs, radio stations, and movies to a playlist. To access the playlist, click on the *Playlist* button in the main interface.

The default playlist view.

#### Additional Sources

In addition to audio and video files, you can play other formats. The additional formats supported by VLC media player are described in the following sections:

-   **Podcasts** - Podcast (Personal On Demand broadCASTING) is a series of audio or video digital media files which is distributed over the Internet and downloaded to media players. If consumers subscribe to Podcasts, whenever new content is added the content gets automatically added to the playlist. You can customise Podcasts. To add a Podcast URL

1.  Select the  *Playlist* button.
2.  Click on the *Internet* button to select it in the left pane. The *Podcasts* menu item will appear under *Internet*.
3.  Select a Podcast stream in the main dialog box. Then right-click the stream and select *Play* from the popup menu.

-   **SAP Announcements** – Helps to advertise your stream over the network.

To play a SAP announcement:

1.  Select the  *Playlist* button.
2.  Click on the *Local Area Network* to select it in the left pane. The *Network Streams (SAP)* menu item will appear under *Local Area Network*.
3.  Select an SAP announcement and right-click. Select *Play* from the popup menu.

-   **Shoutcast Radio Listings** – Shoutcast is a server for streaming the media developed by Nullsoft. Digital audio content can be broadcast from and to media players, and this helps individuals to create Internet radio networks. Using VLC media player, you can listen to your favourite radio stations and you can also create bookmarks to listen to these radio stations in future.

To customize a Shoutcast radio listing:

1.  Select the  *Playlist* button.
2.  Select *Icecast Directory* under *Internet* in the *Playlist* menu. A list of radio stations appears in the right hand panel. If nothing appears in the right hand panel try double-clicking the *Shoutcast Radio* option and wait. It may take a few minutes the first time. After a while, the right hand panel displays a list of titles.

1.  Scroll down and select a radio station.
2.  Right-click on a radio station and:
    1.  Select *Play* if you want to listen to the radio station.
    2.  Select *Remove Selected* if you want to delete the radio station.
    3.  Select the *Stream* option. The Stream output dialog box is displayed. Refer to the Specifying Specifying Streaming options section for more details. Modify the required parameters and click on the *Stream* button to stream the media file.
    4.  Click to select a title in the Playlist dialog box and right-click. Select *Save* from the popup menu. The Stream Output dialog box is displayed. Select the required options and click on the *Save* button in the Stream Output dialog box. Refer to the Specifying Streaming options section for more details.
    5.  Select the Information option. The Media Information dialog box is displayed with details of the media being played.
    6.  Select *Title* to alphabetically sort the radio stations.
    7.  Click to select a title in the Playlist dialog box and right-click. Select *Open Folder* from the popup menu. A folder is opened to show all sub nodes within a title.
    8.  Select *Add Node* to add a node.
    9.  Click to select a title in the Playlist dialog box and right-click. Select *Information* from the popup menu to view the details of the selected title. Refer to the *Media Information* section for more details on options.

-   *Shoutcast TV stream* – You can watch streaming TV using the VLC media player. Shoutcast TV stream refers to a stream transmitted by Nullsoft. The procedure of customising the TV stream and the options are similar to that of the Shoutcast Radio.
-   *Freebox TV listing* – Refers to television service over ADSL accessible by Freebox Free Zone unbundled.

*Note:* You should be connected to the Internet to access these streams.

#### Add Media Files to Playlist

You can add several media files to a playlist. The media files could be selected from the media library, additional sources, and some other source.

To add files to a playlist:

1.  Select the  *Playlist* button.
2.  Right-click on the dialog box and click and a short list appears with two options: Add file and Add directory.
    1.  Select *Add file* to add a file to the playlist.
    2.  Select *Add directory* to add a directory containing media files to the playlist.
3.  Click on the  **Random** icon. This icon toggles between *Random* and *Random Off*. Click on  to play files at random. Click on  and the files are played in an order.
4.  Click on the  *Repeat* icon. This icon toggles between *Repeat One* and *Repeat All*. If you want to listen to a track several times, click on  icon. If you want to listen to all tracks, click on  again.
5.  To search for a media file, enter the name in the *Search* box. To search for media files with certain names or formats, enter a word or phrase in the *Search* box. All files with the specified name are listed.
6.  Click on the  icon. This icon is used to skip to the current item when you have a very long list.
7.  Click the *Remove Selected* button to clear a track from the playlist.

#### Load Playlist

This option is used to add a playlist created in some other media player. You can load playlists of the *.xspf, .asx, .b4s* and *.m3u* formats. To load a playlist:

1.  Select the *Open* option from the *Media* menu. The *Open file* dialog box is displayed.
2.  In the bottom right, change the format to *Playlist Files* in the selector.
3.  Locate a playlist file and click on *Open*.

The selected playlist is added in the current playlist dialog box.

#### Save Playlist

You can save playlists using the VLC media player in format of your choice. To save a playlist:

1.  Create a playlist. Refer to Add Media Files to Playlist for creating a playlist.
2.  Select *Save Playlist to File* from the *Media* menu. The *Choose a filename to save playlist* dialog box is displayed.
3.  Select a name for the playlist.
4.  Select a format in which the playlist must be saved from the *Files of type* list. The Files of type list contains the *.xspf* and *.m3u* formats.
5.  Click on *Save* to save the playlist in the selected format.

#### Play a file

To play a file, open the Media menu, and select the Open File menu item. An Open File dialog box will appear. Select the file you want to open, and click Open. VLC will start playing the selected file. An alternative is to drag 'and' drop your file onto the VLC main interface or playlist window from the file explorer (Finder on MacOS X).

VLC 0.9.8a version Windows XP mode

The File menu - MacOS X interface**- needs verifying for 0.9**

The Open file dialog - wxWidgets interface

(FIXME need 0.9 screenshot for MacOS) The Open file dialog - MacOS X interface

#### Naming Files

You can change the original file name to one you would like before adding the file to VLC. When adding files from the menu bar, the new file name will show in the playlist. However, when dropping the file using the "add/drop" option, VLC may not recognize the name change depending on the file type. In that case, you can right click the header of the playlist column and select "URL," you will then see the original file path for the file.

#### Sorting

In the wxWidgets interface, *Sort* allows you to sort the playlist according to several criteria, or to shuffle it. You can also sort by clicking the header of the column.

In the MacOS X interface, sorting can be done by clicking the header of the column matching the criteria you want to use for sorting.

#### Playlist modes

The playlist supports several playback modes.

In the wxWidgets interface, the toolbar contains three playlist mode buttons. They allow you to enable random mode, to repeat the whole playlist or to repeat one item.

In the MacOS X interface, random mode can be enabled by selecting the *Random* box. A drop down menu allows you to enable playlist and item repeat modes.

#### Misc

##### Search

You also have a search tool. Enter a search string and hit search. The next item to match the string will be highlighted. Keep hitting Search to cycle between all matching items.

##### Moving items

In the wxWidgets interface, the *Up* and *Down* buttons at the bottom of the playlist window allow you to move an item. Select an item and use these buttons to move it.

In the MacOS X interface, you can easily move an item with the mouse, using drag-and-drop.

##### Contextual menu

By right-clicking or control-clicking an item, a contextual menu will appear, giving access to a number of functions (for example, play the item, disable it, delete it, or get info on it).

##### Example finding a Shoutcast radio stream

This example was verified as working on 15 October 2008, using VLC 0.9.4 under Windows Vista. *This needs reproducing by other people on other versions and other operating systems.*

1\. Ensure your firewall is set to allow the VideoLan program to make outgoing connections.

2\. Click *Tools* then *Preferences*, click Interface and then click All under "Show settings". Then click the "-" next to "Playlist" in order to show the "Services discovery" submenu. If the shoutcast radio listings box is empty, click it so that a check-mark appears. The text field underneath should now show the word "shout". Click the Save button to save and close the Preferences window:

3\. Restart VLC media player to make it take notice of the changed preferences.

4\. On the VLC interface click *Playlist*, then click *Show Playlist*. Select the "Shoutcast Radio" in the left hand panel. If nothing appears in the righthand panel, try double-clicking "Shoutcast Radio" and waiting, it may take a few minutes the first time. After a while the righthand panel displays a long list of titles.

5\. Scroll down the radio stations in the right-hand panel and select one. Click the mouse right button and click the "Play" item.

6\. It may take some time for the connection to the radio station to establish (and it may fail if the station's outgoing streams are all occupied). When it does connect, VLC should start playing the audio stream from the station:

##### Example playing a known Shoutcast radio stream

Go to 0 and search for a radio station of your choice. On Windows, right-click your mouse over Shoutcast's "Tunein" button and click "Save Link As..." to save the playlist on your computer. Remember where you saved the playlist, rename it to something that makes sense.

At any time later, you can use VLC to open the saved playlist and listen to that radio station.

For example, to find a BBC World Service radio stream, use a browser to go to: 0

One of the stations listed may be playing the World Service, if so move your mouse over the "TUNEIN!" webicon and click the right mouse button and click "Save Link As...", as described above.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use / Snapshots {#play-howto-basic-use-snapshots}

There are two ways to take snapshots (i.e., screenshots or frame grabs) with VLC:

1.  Open the *Video* menu, and select the *Take Snapshot* menu item.
2.  Press the snapshot hotkey
    -   Linux / Unix / Windows (Qt Interface): Shift+s
    -   Mac OS X: Command+Alt+s

When a snapshot is taken, it will briefly preview as a thumbnail with its filename and then fade away.

To change the hotkey, go to Tools → Preferences. If "Show settings" is set to Simple, click Hotkeys; if "Show settings" is set to All, navigate to Interface → Hotkeys settings. Set the hotkey for Take video snapshot.

#### Snapshot location, format and name

The snapshot location depends upon your operating system:

-   Windows XP: "%HOMEPATH%\\My Pictures\\"
-   Windows Vista, 7, 8, and 10: "%HOMEPATH%\\Pictures\\"
-   Linux / Unix: \~/Pictures
-   macOS: Desktop/

##### Configuring snapshot options under Windows:

The location, format and name of snapshots may be changed in the *Preferences* menu item in the *Tools* tab, subsection *Video*.

The default format for snapshots is PNG, but this may be changed to JPEG. Also, the default name for snapshots is *vlcsnap-* followed by a timestamp that is *not* the time of the frame in the video you're viewing, but rather the current date and time—as in 2014-01-16-14h57m19s163.

Also, you may substitute other text for *vlcsnap-* in the *Video snapshot file prefix* and you may choose to have snapshots numbered sequentially (i.e., 000001, 000002, 000003, and so on) instead of with a timestamp.

As of version 0.9.0, you may even use variables in the text used for the filename. For example, *\$T* (must be upper case) will insert the video's time code into the file name. If you were to change the prefix to *Friends-\$T-* while watching a DVD of *Friends*, then the snapshot filenames would look something like this: *Friends-00_05_21-2014-01-16-14h57m19s163.png*. This indicates a snapshot taken at 5 minutes and 21 seconds into the video; and it was taken on this day at this time: *2014-01-16-14h57m19s163*.

For a shorter file name, check the "Sequential numbering" option in the configuration box (below). Instead of numbers like *2014-01-16-14h57m19s163*, VLC will simply insert the count of snapshots for that session—for example, *00004*. Thus, in the example above, a snapshot with sequential numbering would look like this: *Friends-00_05_21-000001.png*

For a full list of variables, please see Documentation:Play HowTo/Format String.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use / Subtitles {#play-howto-basic-use-subtitles}

VLC supports many kinds of subtitles.

#### Media with included subtitles

Many types of media can have embedded subtitles. VLC can read subtitles for the following media formats:

-   DVD
-   SVCD
-   OGM files
-   Matroska (MKV) files

Subtitles are enabled by default in VLC media player. To disable them, go to the *Video* menu, and to *Subtitles track*. All available subtitles tracks will be listed. Select "Disable" to turn off the subtitles. Depending on the media, a description (language, for example) might be available for the track.

To disable subtitles by default, select "Preferences", then "Show All". Select "Input/Codecs". On the "Subtitle Track ID" selection window, change the value to "-1". (NOTE: Changing the value in the "Subtitle Track" menu will not disable the subtitle file.) In the case of multiple subtitle tracks, a value of "0" will enable subtitle track 1, a value of "1" will enable subtitle track 2, and so on.

VLC under Linux:

VLC under OSX:

VLC under Windows:

DVD and SVCD subtitles are merely images, so you won't be able to change anything for them. OGM and Matroska subtitles are rendered text, so you will be able to change several options.

Text rendering options can be changed in the *Preferences* in the *Tools* tab. To adjust the font preference check the *All* bullet in the *Show Settings* box, and then click *Subtitles/OSD*. You can then set the font and its size under *Text Renderer*. For the font, you have to select a font file. In Windows, they can be found in *C:\\Windows\\Fonts*. Under MacOS X, they are in */System/Library/Fonts*. Sizes can be set either relatively or as a number of pixels.

Subtitle text rendering preferences under Windows, VLC 1.1.5

You need to restart your stream for the font modifications to take effect.

#### Subtitles files

While modern file formats like Matroska or OGM can handle subtitles directly, older formats like AVI can't. Therefore, a number of subtitles files formats have been created. You need two files: the video file and the subtitles file that only contains the text of the subtitles and timestamps.

VLC can handle these types of subtitles files:

-   MicroDVD
-   SubRIP
-   SubViewer
-   SSA
-   Sami
-   Vobsub (this one is quite special: it is not made from text but from images, which means that you can't change the fonts)

To open a subtitles file, use the Advanced Open dialog box (Menu File, Open file). Select your file by clicking on the *Browse* button. Then, check the *Subtitle options* checkbox and click on the "Settings" button.

You can then select the subtitles file by clicking the *Browse* button. You can also set a few options like character encoding, alignment and size.

An alternative is loading subtitles from the *Subtitles Track* menu item under the *Video* tab.

Note: For Vobsub subtitles, you need to select the **.idx** file, not the **.sub** file. Encoding, alignment and size won't have any effect for Vobsub subtitles.

Font can be changed as explained in the previous section.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use / Video {#play-howto-basic-use-video}

You can play video files, video clips and other video media using the VLC media player. You can resize, change aspect ratio, crop videos, load subtitles, deinterlace, save snapshots, and convert videos to DirectX wallpaper.

Video tracks of the *.asf, .avi, .divx, .dv, .mxf, .ogg, .gm, .ps, .ts, .vob,* and *.wmv* formats are supported.

#### Playing a Video Track

There are two main ways to open and play a video track:

1.  Select *Open File* from the *Media* menu.

     2. Select a video track and double-click it or click the *Open* button.

The selected track will play.

#### Loading Subtitle Tracks

A subtitle is a textual version of a movie’s dialogue. Subtitles are helpful if you are viewing a movie that contains foreign language(s). You can load subtitles for video tracks. Subtitles of the formats *.cdg, .idx, .srt, .sub, .utf, .ass, .ssa, .aqt, .jss, .psb, .rt* and *smi* are supported.

VLC can read subtitles for the media formats such as *DVD*, *SVCD*, *OGM* files, and *Matroska (MKV)* files.

To enable the subtitle for a track:

1.  Select *Open File* under the *Subtitle* menu item from the *Video* menu. The *Open Subtitles File* dialog box is displayed.  
2.  Locate the file which contains the subtitle and click on *Open*. The subtitles are displayed.

For more details, see [Documentation:Subtitles](#subtitles).

#### Full Screen

This option is useful if you want to watch the video in the full screen mode.

1.  Select *Full Screen* from the *Video* menu. The video will then occupy the entire screen.
2.  To return to the original mode, press *Esc* on the keyboard or right-click the mouse and select the *Leave Full Screen* option. The video will then return to its original mode.

Note: When you switch to full screen, the controls may appear for a short period of time. To restore the controls after they disappear, move the mouse or press any key on the keyboard.

#### Always on Top

This option is useful if you want the VLC media player to remain on the top of the screen always when other applications or files are open.

1.  To make the VLC media player appear on top of the screen, select *Always on Top* from the *Video* menu. 
2.  If you do not want VLC to appear on the top of the screen, select the *Always on Top* option from the *Video* menu and manually minimise the VLC application.

#### DirectX Wallpaper

This option is useful if you want to display the video which is being played as your desktop wallpaper.

To view the current video file as wallpaper

1.  Select *Advanced File Open* from the *Media* menu. The *Open Media* dialog box is displayed. 
2.  Select a file and click  *Play*.
3.  Select *DirectX Wallpaper* from the *Video* menu.

The wallpaper mode will then display the video as the desktop background.

Note: that this feature works only if you deactivate the overlay under Windows XP.

#### Snapshot

This option is useful if you want to capture a portion of the video as an image.

1.  Select *Advanced File Open* from the *Media* menu. The Open dialog box is displayed.
2.  Select a file and click  *Play*.
3.  To capture an image from the video, select *Snapshot* from the *Video* menu.

The image is captured in the *.png* picture format and is saved in the *C:\\My Pictures* folder by default (*C:\\Users\\**Username**\\Pictures*).

#### Zoom

You can enlarge videos in different sizes. This option is useful if you want to change the size of a video track which is being played. The supported sizes are *1:4 Quarter, 1:2 Half, 1:1 Original (default)* and *2:1 Double*.

To view a video in a particular dimension, select a dimension from *Zoom* in the *Video* menu. The track is then resized based on the selected zoom ratio.

#### Aspect Ratio

Aspect ratio refers to the width of a picture in relation to its height. For example, the ratio 4:3 means four units wide to three units high. VLC provides a list of aspect ratio values which are *Default, 1:1, 4:3, 16:9, 16:10, 2.21:1, 2.35:1, 2.39:1* and *5:4*.

To select an aspect ratio, select a value from *Aspect Ratio* in the *Video* menu. The video is then adjusted based on the selected ratio.

#### Crop

This option is helpful if you want to capture a small portion of a video as an image. This also helps crop the black bars of the top and bottom of a video.

The cropping values that are supported are *Default, 16:10, 16:9, 1.85:1, 2.21:1, 2.35:1, 2.39:1, 5:3, 4:3, 5:4,* and *1:1*.

To crop a video that is played, select a value from *Crop* in the *Video* menu. The video is then cropped based on the selected value.

#### Deinterlace

Deinterlace refers to a process where interlaced video signals are converted into non-interlaced signals. VLC provides the *Discard, Blend, Mean, Bob, Linear, X, Yadif and Yadif (2x)* deinterlacement modes.

1.  Select *Deinterlace* from the *Video* menu and choose the appropriate setting.
2.  To change the deinterlacement mode select 'Deinterlace mode' is the *Video Menu*
3.  Select a mode and observe the change in the video being played.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use / Video and Audio Filters {#play-howto-basic-use-video-and-audio-filters}

This page is outdated and information might be incorrect.

VLC includes a system of *filters* that allow you to modify the audio and video.

#### Deinterlacement and Post Processing

VLC is able to deinterlace a video stream using different deinterlacement methods. Deinterlacement can be enabled in the *Video* menu, *Deinterlacement* menu item. The *Blend* methods gives the best results in most cases. The *discard* method is a less resource consuming alternative, although its results may be slightly compromised.

On some particular streams (MPEG 4, DivX, Xvid, Sorenson, etc.), some additional image filtering can be applied to the video before display, improving its quality in some cases. This can be enabled by using the *Post processing* menu item in *Video*. Different levels of post processing can be chosen here. A higher level means more filtering.

#### Video filters

##### Summary

VLC features several filters able to change the video (distortion, brightness adjustment, motion blurring, etc.).

In Windows and Linux, the user must go to the *Effects and Filters* in the *Tools* menu item. A dialogue box entitled "Adjustments and Effects" will appear.

In macOS you can enable these filters through the *Extended Controls panel*. Click on the triangle next to *Video filters* to select your filters or expand the *Adjust Image* section to change the contrast, hue, etc.

iOS:

Example of combined effects on a video:

##### Rotate

You can easily rotate a video. Open the *Effects and Filters* dialog, in the *Tools menu*

Select the *Video Effects* tab, then the *Geometry* one.

Check the *Transform* checkbox to use rotation presets (90°, 180°, 270°) or check the *Rotate* checkbox to manually select the angle you wish to apply.

#### Audio filters

##### Equalizer

[Wikipedia](http://en.wikipedia.org/wiki/Main_Page) has information on this entry:

***[Equalization (audio)](http://en.wikipedia.org/wiki/Equalization_(audio) "wikipedia:Equalization (audio)")***

VLC features a 10-band graphical equalizer, a device used to alter the relative frequencies of audio (e.g. for a bass boost). You can display it by activating the advanced GUI on wxWidgets or by clicking the *Equalizer* button on the macOS interface. The following image is the interface of the audio equalizer in the Windows and GNU/Linux interface.

The equalizer in the macOS interface

Presets are available in all of these dialog boxes.

##### Other audio filters

At the moment, VLC features two other audio filters: a volume normalizer and a filter providing sound spatialization with a headphone. They can be enabled in the *Effects and Filters* menu item in the *Tools* tab of the Windows and GNU/Linux interface and in the Audio section of the Extended Controls panel of the macOS interface.

For better control, you need to go to the preferences. To select the filters to be enabled, go to *Audio*, then to *Filters*. In the "audio filters" box, enter the names of the filters to enable, separated by commas. Valid names are "equalizer", "normvol" and "headphone".

If you want to tune the behavior of these filters, go to *Audio, Filters, \[your filter\]*. The equalizer and headphone filters can be tuned.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use / VLC 1.2 Interface on Ubuntu {#play-howto-basic-use-vlc-1-2-interface-on-ubuntu}

** VLC 1.2 Ubuntu (Linux) Interface Documentation**

All the button functions in Vlc 1.2 Ubuntu (Linux) version

Opening a file (Location--\> Media/Advanced open file)

Media Information (Location--\> Tools/Media Information)

Playlist Control (Location--\> View/Playlist)

Add Adjustments and Effects to the playing Video or Audio (Location--\> Effects and Filters)

Control Vlc rapidly fast with Shortcut keys (Location--\> Toos/Preferences/Hotkeys)

Visualization (Location--\> Audio/Visualization)

Skin vlc of your own choice (Location--\> Tools/Preferences/Interface/)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Basic Use / VLC 1.2 Interface on Windows 7 {#play-howto-basic-use-vlc-1-2-interface-on-windows-7}

VLC 1.2 Windows 7 Interface Documentation

All buttons functions in VLC 1.2 Windows 7 version

All buttons functions in fullscreen

Opening VLC for the first time

How to open a file(click Open, select a file and open it by double clicking on the file or click once on it and press open)

Another way of doing it

or

### Building Lua Playlist Scripts {#play-howto-building-lua-playlist-scripts}

This page contains example code.

#### Introduction

Starting with version 0.9.0, VLC gives you the possible to implement your own playlist loading modules easily. Such modules can do stuff like:

-   *URL translation*: You give it the youtube webpage URL and VLC starts playing the corresponding video;
-   *Text playlist parsing*: You use some custom text playlist format.

Lua playlist scripts shipped with VLC are stored in the following directory:

-   C:\\Program Files\\VideoLAN\\VLC\\lua\\playlist\\ on Windows;
-   VLC.app/Contents/MacOS/share/lua/playlist/ on Mac OS X;
-   /usr/share/vlc/lua/playlist/ on Linux.

You can add your own Lua playlist scripts in this directory or in your VLC's preferences folder "lua/playlist" subdirectory on Windows or Mac OS X and in your local VLC shared data folder on Linux (\~/.local/share/vlc/lua/playlist).

#### Simple Examples

##### URL translation

###### What we want to do

Let's say that we want VLC to open Google Video links automatically. Google Video pages have URLs like

    0

to the URL of the corresponding Google Video Playlist file which is

    0

###### Probe

The Lua script is going to be made of 2 functions. The first one is the **probe()** function. This function tells VLC if the Lua script should be used (function returns true) or not (function returns false). Here it would look like:

     function probe()
         if vlc.access ~= "http"
         then
             return false
         end
         if string.match( vlc.path, "video.google.com/videoplay" )
         then
             return true
         else
             return false
         end
     end

Lets analyse that function step by step. First we check that VLC is using HTTP. If it isn't (for example it's reading a file off your hard drive or a DVD), we don't want to trigger the Google Video URL translation. The **vlc** Lua object provides **vlc.access** which should be equal to string **"http"**. Then we check that the URL is a Google Video page. This is easily done by trying to find the **"video.google.com/videoplay"** string in **vlc.path**.

Note that the same function can be written as

     function probe()
         return vlc.access == "http" and string.match( vlc.path, "video.google.com/videoplay" )
     end

###### Parse

If the **probe()** function returns true, VLC will use the **parse()** to ask for the new playlist item(s) which need to be added.

     function parse()
         item = {}
         item.path = "http://" .. string.gsub( vlc.path, "videoplay", "videogvp" )
         item.name = "Some Google video playlist"
         return { item }
     end

We create a new playlist item (a Lua table). We set the item's path to the appropriate string: 1/ we prepend "http://" since the **vlc.path** string doesn't include that part of the original URL and 2/ we replace "videoplay" by "videogvp".

We also set the new playlist item's name to "Some Google video playlist".

We then return the new playlist to VLC (a Lua table which basically represents a list of items).

A shorter version would be:

     function parse()
         return { { path = "http://" .. string.gsub( vlc.path, "videoplay", "videogvp" ); name = "Some Google video playlist" } }
     end

###### Saving that to a file

To make that script available to VLC, simply create a new something.lua file in one of the directories listed in the introduction (You should also remove the googlevideo.lua file shipped with VLC to make sure that it isn't used instead of your new script).

##### Text Playlist Parsing

In the previous example we translated a URL to another URL. This new URL redirects VLC to a Google Video Playlist which is basically a text file. This file needs to be read to get the final video's true URL. Here's what such a file would look like if you were to open it with a text editor:

    # download the free Google Video Player from 0
    gvp_version:1.1
    url:0
    docid:-5784010886294950089
    duration:72640
    title:Penguins go for a stroll
    description:African penguins caught a breath of fresh air as they were out for a stroll through Tokyo's aquarium. Ikebukuro Sunshine Aquarium offer spectators a chance to get a closer encounter with penguins as they are taken for a walk through the aquariums compound.
    description:
    description:During the parade the penguins were separated from spectators, mainly children, by portable fences on wheels which were pushed by the zookeepers. The fence ensures they don't run away and also prevents them from biting spectators.
    description:
    description:"I came up with this idea to let people get a close look at penguins walking." said Keeper Masahiro Tomiyama, adding only penguins born and raised at the aquarium can walk outside their cage without feeling stressed by all the attention.
    description:
    description:Reuters 16094/06
    description:Keywords: animals, cute, sweet, funny, ITN Source.

###### Probe

The first thing we need to worry about is making sure that the file we're playing is a Google Video Playlist (GVP) file. We could rely on the URL, but that would prevent playing GVP files from our hard drive. Fortunately, the file's contents, especially the "gvp_version:" string seem specific to GVP files. We'll thus try reading a bunch of characters from the beginning of the file and look for the "gvp_version:" string.

     function probe()
         return string.match( vlc.peek( 512 ), "gvp_version:" )
     end

The **vlc.peek(** *\* **)** function asks VLC to return the first \ characters. If the string "gvp_version:" isn't found in a file's first 512 characters, we're almost 100% sure that it's not a valid GVP file.

###### Parse

We now need to read information from the file to create our new playlist item. A simple version would only fetch the URL:

     function parse()
         item = {}
         while true
         do
             line = vlc.readline()
             if not line
             then
                 break
             end
             if string.match( line, "^url:" )
             then
                 item.path = string.sub( line, 5 )
             end
         end
         return { item }
     end

We read all the file's lines using the **vlc.readline()** function. If we encounter the line starting with **"url:"** (*string.match( line, "url:" )* would match lines containing "url:", while *string.match( line, "\^url:" )* only matches those starting with "url:"), we use that as our new item's path.

If vlc.readline() returns nil, this means that we've finished reading the file so we break out of the while loop and return our new playlist.

This simple **parse()** function unfortunately discards all the other useful meta information available in the GVP file. Lets try to use it:

     function parse()
         item = {}
         while true
         do
             line = vlc.readline()
             if not line
             then
                 break
             end
             if string.match( line, "^url:" )
             then
                 item.path = string.sub( line, 5 )
             elseif string.match( line, "^title:" )
             then
                 item.name = string.sub( line, 7 )
             elseif string.match( line, "^duration:" )
             then
                 item.duration = string.sub( line, 10 ) / 1000
             elseif string.match( line, "^description:" )
             then
                 if item.description
                 then
                     item.description = item.description .. "\n" .. string.sub( line, 13 )
                 else
                     item.description = string.sub( line, 13 )
                 end
             end
         end
         return { item }
     end

Getting the video's title works like the URL parameter. The duration is a bit more tricky. GVP uses times in milliseconds while VLC wants a time in seconds. We thus have to divide it by 1000. For the description, it works like the URL and title parameters except that a GVP file can have more than one description parameter. If we read more than one parameter we thus concatenate the different description lines.

#### Reference

Information about the VLC Lua scripts is available in your VLC installation in the lua subdirectory.

-   [Global VLC Lua README](https://code.videolan.org/videolan/vlc/-/blob/master/share/lua/README.txt)
-   [Playlist script specific README](https://code.videolan.org/videolan/vlc/-/blob/master/share/lua/playlist/README.txt)

#### Scripts shipped with VLC

Scripts for popular playlist formats and video websites are available in the default VLC installer:

-   [dailymotion.lua](https://code.videolan.org/videolan/vlc/-/blob/master/share/lua/playlist/dailymotion.lua): URL translation for Daily Motion video pages;
-   [metacafe.lua](https://code.videolan.org/videolan/vlc/-/blob/master/share/lua/playlist/metacafe.lua): URL translation for metacafe video pages and flash player;
-   [vimeo.lua](https://code.videolan.org/videolan/vlc/-/blob/master/share/lua/playlist/vimeo.lua): URL translation for Vimeo video pages;
-   [youtube.lua](https://code.videolan.org/videolan/vlc/-/blob/master/share/lua/playlist/youtube.lua): URL translation for YouTube video pages and flash player (including fullscreen video URLs);

#### Useful links

-   [Lua 5.1 reference manual](http://www.lua.org/manual/5.1/)
-   [Lua tutorials](http://lua-users.org/wiki/TutorialDirectory)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Building Pages for the HTTP Interface {#play-howto-building-pages-for-the-http-interface}

**This page is obsolete and kept only for historical interest.** It may document features that are obsolete, superseded, or irrelevant. Do not rely on the information here being up-to-date.

#### Introduction

This appendix describes the language used for writing dynamic web pages for the HTTP interface.

Pages must be placed in the 0 folder in either VLC's folder (Windows, Mac) or 1 or 2 (or wherever vlc's shared files are installed).

Some files are handled a bit specially:

-   Files beginning with '.' are not exported.
-   A '.access' file will be opened and the http interface will expect to find at the first line a login/password (written as login:password). This login/password will be used to protect all files in this directory. Be careful that only files in this directory will be protected. (sub-directories won't be protected.)
-   A '.hosts' file will be opened and the http interface will expect to find a list of network/mask pairs separated by new line, for instance 192.168.0.0/255.255.255.0. If this file is present, then the default behaviour is to deny access from hosts which don't match any of the network/mask pairs to all the files of the directory. If the file is not present, then any host has access to the files of the directory. Be careful that only files in this directory will be protected. (sub-directories won't be protected.)
-   The file \/index.html will be exported as \ and \/ and not as index.html.

The MIME type is set by looking at the file extension and cannot be specified nor modified for a specific file. Unknown extensions will have "application/octet-stream" as MIME type.

You should avoid exporting big files. Each file is indeed first loaded into the memory before being sent to the client, so please be careful.

#### VLC macros

Each time a .html/.htm page is requested, it is parsed by VLC before being sent. The parser searches for the VLC macros, and executes or substitutes them. Moreover, URL arguments received by the GET method can be interpreted.

A VLC macro looks like:


"id" is the only mandatory field, param1 and param2 may or may not be present and depend on the value of "id".

You should take care that you **have to** respect this syntax, VLC won't like invalid syntax. (It could easily leads to crashes).

Examples:

Correct:


Incorrect:

      <!--(missing tag ending)-->
      <!--(missing "" )-->

Valid macros are:

-   **control** (1 optional parameter)
-   **include** (1 parameter)
-   **get** (2 parameters)
-   **set** (2 parameters)
-   **rpn** (1 parameter)
-   **if** (1 optional parameter)
-   **else** (no parameter)
-   **end** (no parameter)
-   **value** (1 optional parameter)
-   **foreach** (2 parameters)

For powerful macros, you may use these tools:

-   RPN Evaluator (see part 2)
-   Stacks: The stack is a place where you can push numbers and strings, and then pop them backs. It's used with the little RPN evaluator.
-   Local variables: You can dynamically create new variables and changes their values. Some local variables are predefined:
    -   **url_value**: parameter of the URL
    -   **url_param**: 1 if url_value isn't empty else 0
    -   **version**: the VLC version
    -   **copyright**: the VLC copyright
    -   **vlc_compile_time, vlc_compile_by, vlc_compile_host, vlc_compile_domain, vlc_compiler, vlc_changeset**: information on the VLC build
    -   **stream_position, stream_time, stream_length, stream_state**: information on the currently playing stream
    -   **volume**: current volume setting

Remark: The stacks, and local variables context is reset before the page is executed.

#### The RPN evaluator

RPN means Reverse Polish Notation.

##### RPN Introduction

RPN may look strange but it's a fast and easy way to write expressions. It also avoids the use of 0 and 1.

Instead of writing 0 you just use 1.

The idea behind it is: if we have a number or a string (using ''), push it on the stack. If it is an operator (like 0), pop the arguments from the stack, execute the operators and then push the result onto the stack. The result of the RPN sequence is the value on the top of the stack. A step by step explanation of the sequence **1 2 + 5 \*** is shown below, to illustrate this process:

-   **Stack Contents**: Word Action taken on the stack
-   **empty**: 1 1 is pushed on the stack
-   **1**: 2 2 is pushed onto the stack, 'above' 1
-   **1 \| 2**: + The plus operator results in removal of 1 and 2 from the stack, then write 3 onto the stack
-   **3**: 5 5 is pushed on the stack
-   **3 \| 5**: \* The multiplication operator removes 3 and 5 and writes 15 onto the stack.
-   **15**: Final result.

##### Operators

Notation: ST(1) means the first stack element, ST(2) the second one … and op is the operator.

You have access to :

-   Standard arithmetics operators: **+, -, \*, /, %** these ones push the result of ST(1) op ST(2) onto the stack
-   Binary operators: **!** (push !ST(1)); **\^, &, \|**: push the result ST(1) op ST(2)
-   test: **=, \<, \<=, \>, \>=**: execute ST(1) op ST(2) and push -1 if true else 0
-   string functions:
    -   **strcat**: pushes the result of 'ST(1)ST(2)
    -   **strcmp**: compares ST(1) and ST(2) (0 if equal)
    -   **strncmp**: compares the first ST(1) characters of ST(2) and ST(3) (0 if equal)
    -   **strsub**: extracts characters ST(2) to ST(1) of string ST(3)
    -   **strlen**: pushes the length of ST(1)
    -   **str_replace**: replaces string ST(2) with ST(1) in ST(3)
    -   **url_encode**: encodes non-alphanumeric characters of ST(1) as %XX so that they can be safely passed as GET or POST variables
    -   **url_extract**: performs the reverse operation of url_encode
    -   **addslashes**: protects single quotes (') and double quotes (") of ST(1) with backslashes (\\) so that they can be safely passed to a VLC playlist function
    -   **stripslashes**: performs the reverse operation of addslashes
    -   **htmlspecialchars**: encodes &, ', ", \<, and \> of ST(1) as their &stuff; HTML counterpart, so that they don't interact with HTML tags
    -   **realpath**: parses ST(1) as a filename path, and pushes an absolute path to that file, removing \~ and ../
-   stack manipulation:
    -   **dup**: pops ST(1) and pushes the same string twice
    -   **drop**: pops ST(1) and drops it
    -   **swap**: exchanges ST(1) and ST(2)
    -   **flush**: empties the stack
-   variables manipulation:
    -   **store**: stores ST(2) in a local variable named ST(1)
    -   **value**: pushes the value of the local variable named ST(1)
-   player control:
    -   **vlc_play**: plays the playlist item whose ID is ST(1), and pushes the return value of the play function (0 in case of success); see playlist functions below
    -   **vlc_stop**: stops the playlist
    -   **vlc_pause**: pauses the playlist
    -   **vlc_next**: plays the next playlist item
    -   **vlc_previous**: plays the previous playlist item
    -   **vlc_seek**: seeks the current input to a location defined in ST(1), for instance+3m (minutes), -20s, 45%, 1:12, 1h12m25s
    -   **vlc_var_type**: pushes the type of the variable ST(2) of object ST(1); the type is one of these strings **VLC_VAR_BOOL, VLC_VAR_INTEGER, VLC_VAR_HOTKEY, VLC_VAR_STRING, VLC_VAR_MODULE, VLC_VAR_FILE, VLC_VAR_DIRECTORY, VLC_VAR_VARIABLE, VLC_VAR_FLOAT, UNDEFINED** (no such variable) or **INVALID** (no input stream); the object is one of **VLC_OBJECT_ROOT, VLC_OBJECT_VLC, VLC_OBJECT_INTF, VLC_OBJECT_PLAYLIST, VLC_OBJECT_INPUT, VLC_OBJECT_VOUT, VLC_OBJECT_AOUT** or **VLC_OBJECT_SOUT**
    -   **vlc_var_set**: sets variable ST(2) of object ST(1) to ST(3)
    -   **vlc_var_get**: pushes the value of the variable ST(2) of object ST(1)
    -   **vlc_object_exists**: checks if object ST(1) exists
    -   **vlc_config_type**: pushes the type of the configuration variable ST(1); see **vlc_var_type** for a list of types
    -   **vlc_config_set**: sets configuration variable ST(1) to ST(2)
    -   **vlc_config_get**: pushes the value of the configuration variable ST(1)
    -   **vlc_config_save**: saves the modification made to the configuration variables of module ST(1) to the configuration file (ST(1) may be empty, in which case the whole configuration is saved) and pushes the return status (0 in case of success)
    -   **vlc_config_reset**: resets the configuration file to the default value (use with caution)
    -   **vlc_volume_set**: sets the volume value to ST(1) which can be a raw value between 0 and 1024, or a relative one between 0% and 400%, where 1% is equal to the maximum volume value divided by 400 (thus, the maximum volume value is equal to 400%, that is 1024). If ST(1) begins with a '+' (or '-') operator, the volume is increased (or decreased) by the raw value which follows this operator
    -   **vlc_get_meta**: pushes the value of the meta information named by ST(1) from the stream being played. Available meta names are: "Title" (or "TITLE"), "Author", "Artist" (or "ARTIST"), "Genre" (or "GENRE"), "Copyright", "Album/movie/show title" (or "ALBUM"), "Track number/position in set", "Description", "Rating", "Date", "Setting", "URL", "Language", "Now Playing", "Publisher"
    -   **vlm_command** or **vlm_cmd**: sends the command that is on the stack to the VLM (VideoLan Manager). Since the command can contain more than one component on the stack, it must be ended by an ';' or an empty string pushed on the stack (e.g.: param1="';' 'command' 'my' 'this is' vlm_command"). Once the VLM has executed the command, the return value is assigned to the local variable **vlm_value** and the error string (if available) is assigned to **vlm_error**
    -   **snapshot**: takes a snapshot
-   playlist functions:
    -   **playlist_add**: adds MRL ST(1) to the playlist, with name ST(2) and returns the playlist ID associated to this item; special characters must be escaped with addslashes first; it is very convenient to call 'toto.mpg' playlist_add vlc_play
    -   **playlist_empty**: clears the playlist of all items
    -   **playlist_move**: moves playlist item at position ST(2) to position ST(1)
    -   **playlist_delete**: deletes playlist item ID ST(1)
    -   **playlist_sort**: sorts the playlist using the mode ST(2) and order ST(1). Available order values are 0 (normal order) and 1 (reverse order). Available mode values are 0 (sort by ID), 1 (sort by title), 2 (sort by title, nodes first), 3 (sort by author), 4 (sort by genre), 5 (sort randomly), 6 (sort by duration), 7 (numerically sort by title) and 8 (sort by album)
    -   **services_discovery_add**: adds the service discovery ST(1) to the playlist
    -   **services_discovery_remove**: removes the service discovery ST(1) from the playlist
    -   **services_discovery_is_loaded**: checks if the service discovery ST(1) is loaded in the playlist, and pushes the answer on the stack

#### The macros

##### The *control* macro

**The use of the control macro is now deprecated in favour of the RPN functions above. The documentation is provided here for the maintenance of HTML pages still using this old API. The main problem with this API is that there is no way to retrieve the playlist ID of the last added item.**

When asking for a page, you can give arguments to it through the url. (e.g. using a 0). Ex: *1 The "control" macro tells a page to parse these arguments and to execute the ones that are allowed. param1 of this macro says which commands are allowed. If empty, all commands will be permitted.

Some commands require an argument that must be passed in the URL too.

-   URL commands
    -   Name, Argument, Description
    -   **play**, item (integer), Play the specified playlist item
    -   **stop**, ,Stop
    -   **pause**, Pause
    -   **next**, , Go to next playlist item
    -   **previous**, , Go to previous playlist item
    -   **add**, mrl (string), Add a MRL to the playlist
    -   **delete**, item (integer), Delete the specified playlist item or list of playlist items
    -   **empty**, , Empty the playlist
    -   **close**, id (hexa), Close a specific connection
    -   **shutdown**, , Quit VLC

For example, you can restrict execution of the **shutdown** command to protected page (through a *.access* file), using the control macro in all unprotected pages.

##### The *include* macro

This macro is replaced by the contents of the file param1. If the file contains vlc macros, they are correctly parsed and replaced.

##### The *get* macro

This macro will be replaced by the value of the configuration variable which name is stored in param1 and which type is given by param2.

param1 must be the name of an existing configuration variable. param2 must be the right type of the variable. It can be one of *int*, *float*, or *string*.

Example:

      will be replaced in the output page by the value of sout.

##### The *set* macro

This macro allows to set the value of a configuration variable. The name is given by param1 and the type by param2 (like for get). The value is retrieved from the url using the name given in param1.

For example, if player.html contains


and if you browse at *0*, the 1{.variable} variable will be set to "sout_value". If the URL doesn't contain sout, nothing will be done.

##### The *rpn* macro

This macro allows you to interpret RPN commands. (See II).

##### The *if,else,end* macro

This macro allows you to control the parsing of the HTML page.

If param1 isn't empty, it is first executed with the RPN evaluator. If the first element from the stack is not 0, the test value is true, else false..


         <!-- Never reached -->

         Test succeed: 1 isn't equal to 2


You can also just use "if" and "end".

##### The *value* macro

If param1 isn't empty, it is first executed with the RPN evaluator. The macro is replaced with the value of the first element of the stack.

Note: If the element is the name of a local variable, its value will be displayed instead of its name.

##### The *foreach,end* macro

param1 is the name of the variable that will be used for the loop. param2 is the name of the set to be built:

-   integer: take the first element from the stack to construct a set of integer. The stack element should be a string like: 0 (Ex:1:5:2,6:8:1 will be expanded into 1,3,5,6,7,8)
-   directory: take the first element of the stack as the base directory and construct a set of filename and directly in it. Each element has the following fields:
    -   basename: file/directory name
    -   name: complete file/directory name (including path)
    -   ext: file extension in lowercase
    -   type: "directory" or "file" or "unknown"
    -   size: size of the file
    -   date
-   playlist: set based on the playlist with fields: current is 1 if item is currently selected, 0 else. index is the index value, that can be used by the play or delete control command. name is the

name of the item.

-   "information": Create information for the current playing stream. name is the name of the category, value is its value, info is a new set that can be parsed with a new foreach (subfields of info are name and value).
-   input variables such as "program", "title", "chapter", "audio-es", "video-es" and "spu-es": Create lists for the current playing stream. Every list has the following fields:
    -   name: item name (language for elementary streams, tracks, etc.) to display in places where a human-readable format is preferred
    -   id: item ID to pass to the RPN function vlc_var_set, and returned by vlc_var_get
    -   selected: 1 if the item is selected, 0 otherwise
-   the name of a foreach variable if it's a set of set of value.

&nbsp;


                 :


For more details, have a look at the [share/http](https://git.videolan.org/?p=vlc.git;a=tree;f=share/http;hb=HEAD) directory of the VLC source tree…

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Format String {#play-howto-format-string}

#### Time

Starting with VLC media player 0.9.0, the following options specify a character formatted time string, rather than just a plain character string:

-   --marq-marquee
-   --snapshot-path
-   --snapshot-prefix
-   --sout-file-format
-   --sout-livehttp-index

Time variables are those defined by the **strftime** C function. The following expansions are most common:

-   %Y : year
-   %m : month
-   %d : day
-   %H : hour
-   %M : minute
-   %S : second

For an extensive list have a look at [\[1\]](http://pubs.opengroup.org/onlinepubs/9699919799/functions/strftime.html) (or the strftime manual page on Unix systems).

#### Input meta

VLC-specific meta-data expansions are available for the following options:

-   --input-title-format
-   --snapshot-path (in version 2.2.0 and later)
-   --snapshot-prefix (in version 2.2.0 and later)

The following expansion are performed:

-   \$a : artist
-   \$b : album
-   \$c : copyright
-   \$d : description
-   \$e : encoded by
-   \$f : total decoded frame count (since VLC started)
-   \$g : genre
-   \$l : language
-   \$n : track number
-   \$o : track total
-   \$p : now playing
-   \$r : rating
-   \$s : subtitles language
-   \$t : title
-   \$u : url
-   \$A : date
-   \$B : audio bitrate (in kb/s)
-   \$C : chapter (as in DVD chapter number)
-   \$D : duration
-   \$F : full name with path
-   \$I : title (as in DVD title number)
-   \$L : time left
-   \$N : name (media name as seen in the VLC playlist)
-   \$O : audio language
-   \$P : position (in %)
-   \$R : rate
-   \$S : audio sample rate (in kHz)
-   \$T : time code of the video
-   \$U : publisher
-   \$V : volume
-   \$Z : now playing (artist - title)
-   \$\_ : new line
-   \$\ : \ (for example: \$\$ transforms to \$)

You can insert a space between the \$ sign and the character to tell it to not display anything if the meta data isn't available. For example: 0 instead will display "" while 1 would display "--:--:--", if no time is available. If a time is available, it would display something like "01_22_13" (for a snapshot from one hour, 22 minutes and 13 seconds in a video).

##### Source code

If you want to know how this works, check out [src/text/strings.c](https://git.videolan.org/?p=vlc.git;a=blob;f=src/text/strings.c) (search for 0)

The variable 0{.variable} refers to the leading space feature: [Add option to format strings to prevent displaying dashes if the meta info was unavailable (ie: if time is unavailable, "\$T" will display "--:--:--" while "\$ T" won't display anything). This is of course completely untested :)](https://git.videolan.org/?p=vlc.git;a=commitdiff;h=3cc651e520ccb3106097e3f0167cc8c26f23e36c)

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Introduction to VLC {#play-howto-introduction-to-vlc}

#### Overview of the VideoLAN project

VideoLAN was a complete software solution for video streaming and playback, developed by students of the [Ecole Centrale Paris](http://www.ecp.fr) and developers from all over the world, under the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) (GPL).

Originally VideoLAN was designed to stream MPEG videos on high-bandwidth networks, but VideoLAN's main software, VLC media player, has evolved to become a full-featured, cross-platform media player.

Now the Non-Profit Organisation developing and offering the VLC media player is called: VideoLAN Organisation

More details about the project can be found on the [VideoLAN Web site](https://www.videolan.org/).

#### VLC Media Player

VLC 2.0 default interface, Windows

Originally called *VideoLAN Client*, VLC media player is VideoLAN's main software product.

VLC media player works on many platforms: Linux, Windows, macOS, BeOS, BSD, Solaris, Android, iOS, QNX and many more... It supports the following video and audio formats: MPEG-1, MPEG-2, MPEG-4/DivX, h264, webm, mkv, DVDs, VCDs, Audio CDs, wmv and wma.

It can also play from external sources:

-   Satellite.
-   Cable.
-   Digital TV cards (DVB-S, DVB-T).
-   Several types of network streams: UDP/RTP Unicast, UDP/RTP Multicast, HTTP, RTSP, MMS, etc.
-   Acquisition or encoding cards.
-   Webcams and other devices.

VLC can also be used as a streaming server. This feature is described in the [Streaming HowTo](#streaming-howto).

This guide describes all the playback (client) aspects of VLC media player.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Play HowTo {#play-howto}

This is the user guide for the VLC media player.

#### VLC User Guide

-   Quick start guide: How to start with VLC.
-   [Installation](#installing-vlc): Installation instructions for several systems.
-   [History](#history): Overview and history of the VideoLAN project.

#### Usage

-   [Interface](#interface): The main interface of VLC media player.
    -   OSX Interface
    -   Windows/Linux Interface
-   [Open Media](#open-media): Open every media you want, the way you want.
-   [Audio](#audio): Visualization, selection of devices...
-   [Video](#video): Cropping, snapshots and screenshots...
-   [Playback](#playback): Navigation through media files (e.g. chapters, bookmarks).
-   [Playlist](#playlist): Creating and managing playlists.
-   [Subtitles](#subtitles): Selection of subtitles
-   [Video and Audio Filters](#video-and-audio-filters): Usage of VLC's filters (equalizer, video filters)
-   [Snapshots](#snapshots): How to create snapshots and screenshots.
-   [Hotkeys](#hotkeys): Configuration of VLC's hotkeys
-   Uninstallation: Uninstallation instructions.
-   Troubleshooting: The VLC Support Guide, an informal, step-by-step guide for troubleshooting most common issues with VLC.

#### Advanced Usage

-   [Using VLC inside a webpage](#webplugin): How to create webpages that use the VLC Web plugin.
-   [Command line](#command-line): Main command line instructions.
-   [Alternative Interfaces](#alternative-interfaces) : HTTP interface and other control interface.
-   [Misc](#misc) : Miscellaneous other things.

#### Appendix

-   Building Pages for the HTTP Interface
-   Format String
-   Building Lua Playlist Scripts
-   VLC Use 0.8. (Versions older than 0.9).

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

## Streaming and Converting {#streaming-and-converting}

### Advanced Streaming Using the Command Line {#streaming-howto-advanced-streaming-using-the-command-line}

See also: Category:Stream output

#### Structure of stream output

Stream output is the name of the feature of VLC media player that allows to output any stream read by VLC to a file or as a network stream instead of displaying it. Different kind of processing can be applied to the stream during this process (transcoding, re-scaling, filters, re-muxing…). Stream output includes different modules, each of them having different capabilities. You can *chain* modules to enhance the possibilities.

Here is the list of the modules currently available :

-   **standard** allows to *send* the stream via an *access output* module : for example, UDP, file, HTTP, … You will probably want to use this module at the end of your chains.
-   **transcode** is used to transcode (decode and re-encode the stream using a different codec and/or bitrate) the audio and the video of the input stream. If the input or output access method doesn't allow pace control (network, capture devices), this will be done "[on the fly](http://en.wiktionary.org/wiki/on_the_fly)", in real time. This can require quite a lot of CPU power, depending on the parameters set. Other streams, such as files and disks, are transcoded as fast as the system allows it.
-   **duplicate** allows you to create a second chain, where the stream will be handled in an independent way.
-   **display** allows you to display the input stream, as VLC would normally do. Used with the *duplicate* module, this allows you to monitor the stream while processing it.
-   **rtp** streams over RTP (one UDP port for each elementary stream). This module also allows RTSP support.
-   **es** allows you to make separate Elementary Streams (ES) out of an input stream. This can be used to save audio and video streams to separate files, for instance.
-   **bridge-out** TODO
-   **bridge-in** TODO
-   **mosaic-bridge** TODO

Each of these modules may take options. Here is the syntax that you must use :

    % vlc input_stream --sout "#module1{option1=parameter1{parameter-option1},option2=parameter2}:module2{option1=…,option2=…}:…"

Some of the module options (option1 in the example) have to be set, others are optional. Option parameters (parameter-option1 in the example) are always optional. These option parameters are also often very advanced settings. If you don't understand their description, this certainly means that you don't need them.

You may also use the following syntax :

    % vlc input_stream --sout-module1-option1=… --sout-module1-option2=… --sout-module2-option1=… --sout-module2-option2=… …

For example, to transcode a stream and send it, use :

    % vlc input_stream --sout '#transcode{options}:standard{options}'

In the following documentation, single bullet points represent options and double bullet points represent item options (sub-options) :

-   0
    -   0

#### Description of the modules

##### standard (alias std)

This module saves the stream to a file or sends it over a network, after having muxed it.

The available options are :

###### access

This option allows to set the medium used to save or send the stream. This is a compulsory option. Available options are :

-   **file**: saves the stream to a file.

Use the *append* option to append the stream to an existing file instead of replacing it :

     standard{ … ,access=file{append}, … }

-   **udp**: streams to a UDP unicast or multicast address.
    -   **caching=** to set the time VLC should buffer data before sending it;
    -   **ttl=\** to set the TTL of the sent UDP packets;
    -   **group=\** to sent packets by burst instead of one by one;
    -   **late=** to drop packets that arrive too late at this stage of the chain;
    -   **raw** if you don't want to wait until the MTU is filled before sending the packet.
-   **http**: streams over HTTP.
    -   **user=\** to enable HTTP basic authentication and set the user;
    -   **pwd=\** to set the basic authentication password;
    -   **mime=\** to set the mime type returned by the server.
-   **https**: streams over HTTP, using a secured SSL/TLS connection.
    -   (same as for http option)
    -   **cert=\**to set the certificate to use;
    -   **key=\** to set the private key file the server should use for the TLS connection;
    -   **ca=\** to set the path to the root CA certificates to use for TLS;
    -   **crl=\** to set the revocation certificate to use for the TLS connection.
-   **mmsh**: streams using the Microsoft MMS protocol. This protocol is used as transport method by many Microsoft applications. Note that only a small part of the MMS protocol is supported (MMS encapsulated in HTTP).
    -   (same as for http module)
-   **rtp**: streams over RTP. This can only be used to stream MPEG-TS over plain RTP. Support for this option has been removed in VLC 0.9.0 and later. You should use the **rtp** stream output module instead.
    -   (same as for the **udp** setting)
-   **shout**: sends the stream to a Shoutcast (or Icecast) server.
    -   **mp3=1**: required if your input file uses MP3 audio.
    -   **bitrate=*N***: bitrate of the stream, integer, required for Icecast to display this publicly

###### mux

This option allows you to set the encapsulation method used for the resulting stream. This option has to be set.

Available options are :

-   **ts**: the MPEG-TS muxer. This the standard muxer used to stream MPEG-2. This muxer can be used with any **access** method. Supported codecs are MPEG 1/2/4, MJPEG, H263, H264, I263, WMV 1/2 and Theora for video, MPEG audio, AAC and a52 for the audio stream.
    -   **pid-video=\** to set the PID of the video track;
    -   **pid-audio=\** to set the PID of the audio track;
    -   **pid-spu=\** to set the PID of the subtitle track;
    -   **pid-pmt=\** to set the PID of the PMT (Program Map Table);
    -   **tsid=\** to set the ID of the resulting TS stream;
    -   **shaping=\** to set the minimum interval during which the bitrate of the stream will remain constant, for variable bitrate streams;
    -   **use-key-frames** uses I-frames as limits for the shaping intervals;
    -   **pcr=\** allows to set at which interval Program Clock References will be sent;
    -   **dts-delay=\** allows to delay PTS (Presentation Time Stamps) from the DTS (Decoding Time Stamp) from the given time;
    -   **crypt-audio** allows to enable encryption of the audio track using the CSA algorithm;
    -   **csa-ck=\** allows to set the key used for CSA encryption.
-   **ps**: the MPEG-PS muxer. This the standard muxer for MPEG 2 files (.mpg). It can be used with the file and http output methods. Supported codecs are MPEG 1/2 and MJPEG for video, MPEG audio and a52 for audio streams.
    -   **dst-delay=\**: It allows to delay PTS (Presentation Time Stamps) from the DTS (Decoding Time Stamp) from the given time.
-   **mpeg1**: the standard MPEG 1 muxer. This muxer should be used instead of ps with MPEG 1 video streams, when saved to a file or streamed over HTTP. Supported codecs are MPEG 1 and MPEG audio.
    -   (same as for the PS muxer)
-   **ogg**: the ogg muxer. This is the muxer from the Xiph project. It can be used with the HTTP and file output methods. Supported codecs are MPEG 1/2/4, MJPEG WMV 1/2 and Theora, audio streams can be vorbis, flac, speex, a52 or MPEG audio.
    -   (none)
-   **asf**: the Microsoft ASF muxer. This is the standard muxer used for streaming by Microsoft applications. Is also used as container for WMA audio files. This muxer can be used with the file and HTTP output methods. Supported codecs are MPEG-4, MJPEG, WMV 1/2 for video, MPEG audio, and a52 for audio streams.
    -   **title=\**;
    -   **autor=\**;
    -   **copyright=\**;
    -   **comment=\**;
    -   **rating=\** allow you to set what will be displayed in the according field of the stream comments.
-   **asfh**: this is a special version of the ASF muxer, that should be used for MMSH streaming. MMSH is the only supported output method. Supported codecs are the same as for ASF.
    -   (same as for ASF)
-   **avi**: the Microsoft AVI muxer. This is very common encapsulation format for MPEG-4 files. The only supported output method is file. Supported codecs are MPEG 1/2/4, H263, H264 and I263 for video, MPEG audio and a52 for audio streams.
    -   (none)

The avi muxer in VLC is known to produce corrupt files.

-   **mpjpeg**: the multipart jpeg muxer. This encapsulation format is mostly used on surveillance video cameras with an integrated web server. Such streams are usually embedded in web pages and seen with standard Internet browsers, as they are seen as a succession of jpeg images. The only supported output method is HTTP. The only usable codec is MJPEG. No sound track can be muxed in such streams.
    -   (none)

###### dst

This option allows to give various information about the location where the stream should actually be saved or sent.

Here is the meaning of the **dst** option depending on the parameter used for the **access** option:

-   If the **file** output method is used, **dst** is the path where the file should be saved.
-   If the **udp** or **rtp** output method is used, **dst** is the unicast or multicast destination address – and, optionally – UDP port, in the form **address:port**.
-   If the **http**, **https** or **mmsh** output method is chosen, **dst** is the address, port and path of the local network interface on which the server should listen for requests. If no address is given, VLC will listen on all the network interfaces. These bits of information have to be supplied using the **address:port/path** syntax.

###### sap

Use this option if you want VLC to send SAP (Session Announcement Protocol) announces. SAP is a service discovery protocol, that uses a special multicast address to send a list of available streams on a server.

This option can only be enabled with the **udp** output method.

###### group

This option allows to specify the name of an optional **group** of streams. A VLC used as a client will use this field to classify the stream.

This option uses a private extension of the SAP protocol. VLC will be the only client able to read this field.

This option can only be used if the **sap** option has been enabled.

###### sap-ipv6

Use this option if you want the SAP announces to be sent using the **IPv6** protocol instead of **IPv4**.

This option can only be used if the **sap** option has been enabled.

###### slp

SLP stands for *Service Location Protocol*. It is an alternative to SAP for session announcement. Use this option if you want to send such announcements.

###### name

Use this option to specify the name of the stream that will be sent in SAP and SLP announcements.

This option can only be used if the **sap** or **slp** option has been enabled.

##### display

This module can be used to display the stream. This is particularly useful in a **duplicate** chain, in order to monitor a stream while it is being saved or streamed.

Available options are :

###### novideo

You can use this option to disable video in the displayed stream.

###### noaudio

You can use this option to disable audio in the displayed stream.

###### delay

You can use this option to introduce a delay in the display of the stream. Delay has to be given in ms (milliseconds).

##### rtp

This module can be used to send a stream using the *RTP (Real Time Protocol)* protocol (see RFC 3550).

Although use of *RTSP* is possible using this module, it won't allow you to make *Video On Demand*. Please have a look at the description of the VLM module for that.

The different available options are :

###### dst

This option allows the destination UDP address to be given. This can be the address of a host or a multicast group. This option has to be given, unless the *sdp=rtsp://*option is given (see below). In the latter case, the stream will be sent to the host doing the *RTSP* request.

###### port

This option allows to set the UDP port used to send the first *elementary stream*. This port has to be even. Other streams will be streamed using even ports directly above this one.

###### port-video

This option allows to set the UDP port used to send the first video *elementary stream*. This port has to be even.

###### port-audio

This option allows to set the UDP port used to send the first audio *elementary stream*. This port has to be even.

###### sdp

This option allows to set the way the SDP (Session Description Protocol) file corresponding the the stream should be made available. Options are :

-   **file://\**, to export the SDP as a local file.
-   **http://\**, to make the file available using the integrated HTTP server of VLC.

The *local interface IP* argument is optional. If not given, VLC will listen on all available interfaces.

-   **rtsp://\**, to make the SDP file available using the *RTSP* protocol (see RFC 2326).

The *local interface IP* argument is optional. If not given, VLC will listen on all available interfaces.

-   **sap**, to export the SDP using the SAP (Session Announcement Protocol, see RFC 2974).

###### ttl

This option can be used to set the *TTL* (Time to Live) of the sent UDP packets.

###### mux

This option allows to set the encapsulation method used to send the stream. See **mux** options of the **standard** module for a description of the available method.

Only **ts** is possible for RTP streams. By default, each elementary stream is sent as a separate RTP medium, i.e. no encapsulation is done.

###### rtcp-mux

This option enables RTP/RTCP multiplexing (see draft-ietf-avt-rtp-and-rtcp-mux), i.e. sends and receives RTCP packets on the same port numbers as RTP packets.

By default, RTCP packets are sent and received on the next port.

###### proto

This selects the transport protocol to carry RTP packets.

Possible values include :

-   **dccp**, accept incoming DCCP connections at the specified IP address (dst=),
-   **sctp**, accept SCTP connections at the specified IP address (dst=), *not implemented yet*,
-   **tcp**, accept TCP connections at the specified IP address (dst=) and use RFC 4571 RTP framing, *not implemented yet,*
-   **udp**, send UDP packets to the specified destination (either unicast or multicast); this is the default value,
-   **udplite**, send UDP-Lite packets to the specified destination (either unicast or multicast).

This options uses UDP-Lite instead of UDP as the transport protocol for RTP and RTCP packets.

###### name

This option can be used to set the name that will be displayed on the client receiving the stream.

###### description

This option can be used to give an additional description of the stream.

###### url

This option allows to give the address of a website with additional information about the stream.

###### email

This option allows to give a contact e-mail address.

##### es

The **es** module can be used to separate the different *elementary streams* from a stream, and save each of them in a different file or send it to a separate destination.

The available parameters are :

###### access-video

Use this option to set the medium used to save or send the video *elementary streams*. Possible values and item options are the same as for the **access** option of the **standard** module (see above).

###### access-audio

Use this option to set the medium used to save or send the audio *elementary streams*. Possible values and item options are the same than for the **access** option of the **standard** module (see above).

###### access

This option can be used instead of both **access-video** and **access-audio** options, when they share the same setting.

###### mux-video

Use this option to set the encapsulation method used for the video *elementary streams*. Possible values and item options are the same as for the **mux** option of the **standard** module (see above).

###### mux-audio

Use this option to set the encapsulation method used for the audio *elementary streams*. Possible values and item options are the same than for the **mux** option of the **standard** module (see above).

###### mux

This option can be used instead of both **mux-video** and **mux-audio** options, when they share the same setting.

###### dst-video

Use this option to set the location where the video *elementary streams* should be saved, sent, or made available. The exact meaning of this option depends on the value of the **access-video** option and is the same as for the **url** option of the **standard** module (see above).

If you use the *%n* string in the url field, VLC will replace it by the number of the audio or video track considered. The *%c* string will be replaced by the name (FourCC) of the codec of the track. *%a* prints the access output used and *%m* the muxer used.

###### dst-audio

Use this option to set the location where the audio *elementary streams* should be saved, sent, or made available. The exact meaning of this option depends on the value of the **access-audio** option and is the same as for the **url** option of the **standard** module (see above).

If you use the *%n* string in the url field, VLC will replace it by the number of the audio or video track considered. The *%c* string will be replaced by the name (FourCC) of the codec of the track. *%a* prints the access output used and *%m* the muxer used.

###### dst

This option can be used instead of both **dst-video** and **dst-audio** options, when they share the same setting.

##### transcode

You can use this module to transcode a stream, e.g., to change its codecs or the encoding bitrates. Some additional processing can be done during this process, such as re-scaling, deinterlacing, resampling, etc.

Depending on the bitrate of the original stream and of the options chosen, transcoding can be a very CPU-intensive task. As a consequence, streaming of a real-time transcoded stream can lead to dropped frames or a jerky image and sound in some cases, when running out of resources.

Available options are :

###### vcodec

This option allows to specify the codec the video tracks of the input stream should be transcoded to.

List of available codecs can be found on the [streaming features page](https://www.videolan.org/streaming/features.html).

###### vb

This option allows to set the bitrate of the transcoded video stream, in kbit/s.

###### venc

This allows to set the encoder to use to encode the videos stream. Available options are:

-   **ffmpeg**: this is the libavcodec encoding module. It handles a large variety of different codecs (the list can be found on the [streaming features page](https://www.videolan.org/streaming/features.html).
    -   **keyint=\** allows to set the maximal amount of frames between 2 key frames;
    -   **hurry-up** allows the encoder to decrease the quality of the stream if the CPU can't keep up with the encoding rate;
    -   **interlace** allows to improve the quality of the encoding of interlaced streams;
    -   **noise-reduction=\** enables a noise reduction algorithm (will decrease required bitrate at the cost of details in the image);
    -   **vt=\** allows to set a tolerance for the bitrate of the output video stream;
    -   **bframes=\** allows to set the amount of B-frames between 2 key frames;
    -   **qmin=\** allows to set the minimum quantizer scale;
    -   **qmax=\** allows to set the maximum quantizer scale;
    -   **qscale=\** allows to specify a fixed quantizer scale for VBR encodings;
    -   **i-quant-factor=\** allows to set the quantization factor of I-frames, compared to P-frames;
    -   **hq=\** allows to choose the quality level for the encoding of the motion vectors (arguments are simple, rd or bits, default is simple \*FIXME\*);
    -   **strict=\** allows to force a stricter standard compliance (possible values are -1, 0 and 1, default is 0);
    -   **strict-rc** enables a strict rate control algorithm;
    -   **rc-buffer-size=\** allows to choose the size of the buffer used for rate control (bigger means more efficient rate control);
    -   **rc-buffer-aggressivity=\** allows to set the rate control buffer aggressiveness \*FIXME\*;
    -   **pre-me** allows to enable pre motion estimation;
    -   **mpeg4-matrix** enable use of the MPEG4 quantization matrix with MPEG2 streams, improving quality while keeping compatibility with MPEG2 decoders;
    -   **trellis** enables trellis quantization (better quality, but slower processing).
-   **theora**: The Xiph.org Theora encoder. The module is used to produce theora streams. Theora is a free patent and royalties-free video codec.
    -   **quality=\**. This option allows to create a VBR stream, overriding **vb** setting. the quality level must be an integer between 1 and 10. Higher is better.
-   **x264**. x264 is a free open-source h264 encoder. h264 (or MPEG4-AVC) is a recent high-quality video codec.
    -   **keyint=\** allows to set the maximal amount of frames between 2 key frames;
    -   **idrint=\** allows to set the maximal amount of frames between 2 IDR frames;
    -   **bframes=\** allows to set the amount of B-frames between an I and a P frame;
    -   **qp=\** allows to specify a fixed quantizer (between 1 and 51);
    -   **qp-max=\** allows to set the maximum value for the quantizer;
    -   **qp-min=\** allows to set the minimum value for the quantizer;
    -   **cabac** enables the CABAC algorithm (slower, but enhances quality);
    -   **loopfilter** enables deblocking loop filter;
    -   **analyse** enables the analyze mode;
    -   **frameref=\** allows to set the number of previous frames used as predictors;
    -   **scenecut=\** allows to control how aggressively the encoder should insert extra I-frame, on scene change.

###### fps

This option allows to set the framerate of the transcoded video, in frames per second; reducing the framerate of a video can help decrease its bitrate.

###### deinterlace

This option allows to enable deinterlacing of interlaced video streams before encoding.

###### croptop

This option allows to crop the upper part of the source video while transcoding. The argument is the number of lines the video should be cropped.

###### cropbottom

This option allows to crop the lower part of the source video. The argument is the Y coordinate of the first line to be cropped.

###### cropleft

This option allows to crop the left part of the source video while transcoding. The argument is the number of columns the video should be cropped.

###### cropright

This option allows to crop the right part of the source video. The argument is the X coordinate of the first column to be cropped.

###### scale

This option allows the give the ratio from which the video should be rescaled while being transcoded. This option can be particularly useful to help reduce the bitrate of a stream.

###### width

This option allows you to give the width of the transcoded video, in pixels.

###### height

This option allows you to give the height of the transcoded video, in pixels.

###### acodec

This option allows you to specify the codec the audio tracks of the input stream should be transcoded to.

List of available codecs can be found on the [streaming features page](https://www.videolan.org/streaming/features.html).

###### ab

This option allows to set the bitrate of the transcoded audio stream, in kbit/s.

###### aenc

This allows to set the encoder to use to encode the audio stream. Available options are :

-   **ffmpeg**: this is the libavcodec encoding module. It handles a large variety of different codecs (the list can be found on the [streaming features page](https://www.videolan.org/streaming/features.html)).
-   **vorbis**. This module uses the vorbis encoder from the Xiph.org project. Vorbis is a free, open, license-free lossy audio codec.
    -   **quality=\** allows to use VBR (variable bitrate) encoding instead of the default CBR (constant bitrate), and to set the quality level (between 1 and 10, higher is better);
    -   **max-bitrate=\** allows to set the maximum bitrate, for vbr encoding;
    -   **min-bitrate=\** allows to set the minimum bitrate, for vbr encoding;
    -   **cbr** allows to force cbr encoding.
-   **speex**. This module uses the speex encoder from the Xiph.org project. Speex is a lossy audio codec, best fit for very low bitrates (around 10 kbit/s) and particularly video conferences.

###### samplerate

This option allows to set the sample rate of the transcoded audio stream, in Hz. Reducing the sample rate is a way to lower the bitrate of the resulting audio stream.

###### channels

This option allows to set the number of channels of the resulting audio stream. This is useful for codecs that don't have support for more than 2 channels, or to lower the bitrate of an audio stream.

###### scodec

This option allows to specify subtitle format the subtitles tracks of the input stream should be converted to.

List of available codecs can be found on the [streaming features page](https://www.videolan.org/streaming/features.html).

###### senc

This allows to set the converter to use to encode the subtitle stream.

The only subtitle encoder we have at this time is **dvbsub**.

###### soverlay

This option allows rendering subtitles directly on the video, while transcoding it.

Do not confuse this option with senc/scodec that transcode the subtitles and stream them.

###### sfilter

This option allows to render some images generated by a so-called *subpicture filter* (e.g. a logo, a text string, etc.) on top of the video.

The list of available *subpicture filters* can be found on the [streaming features page](https://www.videolan.org/streaming/features.html). The Item options of this modules can be found using the following command line :

    % vlc -p --advanced

###### threads

This option allows to set the number of computer processing threads that should be used to encode the streams. Increasing this number to the amount of processors on the computer (or twice this number on Intel P4 HT processors) should improve transcoding performance.

###### vfilter

Uses video filter during transcode process. Parameters of vfilter can be found on the [Advanced Use of VLC Filters](#advanced-use-of-vlc).

The example

    vlc input_file --sout="#transcode{vfilter=adjust{gamma=1.5},vcodec=theo,vb=2000,scale=0.67,acodec=vorb,ab=128,channels=2}:standard{access=file,mux=ogg,dst="output_file.ogg"}"

will adjust *input_file* gamma to 1.5, resize the video size (resolution) by 0.67 (e.g. 1080x720 to 720x480), convert video using the Theora codec with bitrate @ 2000 kb/s and audio using the Vorbis codec with bitrate @ 128 kb/s, encapsulate the video and audio to an Ogg container and save it to *output_file.ogg*.

##### duplicate

This module can be used to duplicate the stream, and so process it through several different chains.

Available options are :

###### dst

This option allows to give the chain through which the duplicated stream should be processed.

**dst** options have to be used in the same duplicate block to actually duplicate the stream.

Any of the stream output module described earlier can be used as parameter of this option.

###### select

This options can be used to duplicate only a part *elementary streams* of a complete stream.

Several criteria can be given, by separating each of them with a comma.
For criteria that need a parameter, such as **es** and **program**, you can also specify a range, using the syntax **criteria=num_start-num_end**.

Available parameters are :

-   **program=**: duplicate only *elementary streams* belonging to the selected program (or SID). This option only works with MPEG-TS streams.
-   **noprogram=**: do not duplicate *elementary streams* belonging to the selected program (or PID). This option only works with MPEG-TS streams.
-   **es=**: duplicate only the *elementary stream* with the selected id.
-   **noes=**: do not duplicate the *elementary stream* with the selected id.
-   **video**: duplicate only video *elementary streams*.
-   **novideo**: do not duplicate video *elementary streams*.
-   **audio**: duplicate only audio *elementary streams*.
-   **noaudio**: do not duplicate audio *elementary streams*.
-   **spu**: duplicate only subtitle *elementary streams*.
-   **nospu**: do not duplicate subtitle *elementary streams*.

Example :

    #duplicate{dst=std{…},select="program=100-200,novideo"}

This *duplicate* chain will only output the non video *elementary streams* belonging to the programs which PID are between 100 and 200.

##### Miscellaneous

Here are a few additional global options :

-   **--sout-all**, **--no-sout-all**: Enable streaming of all ES (default enabled). If disabled VLC will only stream one audio ES and one video ES (the first ones). If sout-all remains enabled, all ES (audio, video and SPU) will be streamed.
-   **--sout-keep**, **--no-sout-keep**: Keep sout open (default disabled) : use the same sout instance across the various playlist items, if possible.
-   **--no-sout-audio**: This option disables audio in the output stream.
-   **--no-sout-video**: This option disables video in the output stream.

##### Simplified Syntax

The stream output also offers a simplified syntax, with which you can only you use the **standard** module's main options :

    % vlc input_stream --sout access/mux://url

where **access**, **mux** and **url** are as defined in the options of the **standard** module.

#### Examples

To fully understand the complex syntax of VLC's stream output, please look at the examples in the next section.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Advanced streaming with samples /  multiple files streaming /  using multicast in streaming {#streaming-howto-advanced-streaming-with-samples-multiple-files-streaming-using-multicast-in-streaming}

**This document explains how to stream a file, stream multiple files, use multicast, etc., using the VideoLAN solution. With examples.**

#### UDP streaming examples

Standard UDP streaming:

    % vlc -vvv file:////home/vlc/2007.avi --sout '#std{access=udp,mux=ts,dst=:1234}'

Nothing impossible yet. Streaming a file 2007.avi from /home/vlc/ to udp port 1234.

#### Multicast RTP streaming examples

    % vlc -vvv file:////home/vlc/Jumper.avi --sout '#rtp{access=udp,mux=ts,dst=224.255.1.1,port=1234,sap,group="Video",name=Jumper Movie"}' :sout-all

Hard? No!
This is known key **file**. The key **--sout** starts output as in the UDP sample. Then we set **#rtp** with access type **udp**, muxer **ts**. Then point to multicast IP address **224.255.1.1** with port **1234**. And some keys. We point VLC to do announcements of this stream using **SAP** (see service advertisements protocol), set description of the streaming group to **Video**, and name this stream **'Jumper Movie'** .

#### Multicast RTP streaming with multiple source files (with examples)

*When you start this, you can't stop.*
I spent several hours trying to find this solution. Here it is:

    % vlc -vvv --color -I telnet --telnet-password "i_dont_know_this_password" --vlm-conf=/home/vlc/vlc.streaming.conf

We told that VLC must colorize its output using key **--color**. Then we told VLC to open the telnet server. We must control it, really?! This is the **-I telnet** key. And we the set the password **"i_dont_know_this_password"** to get access to the console. We use the standard VLC telnet port 4212. If you need to change it, use **--telnet-port xxx**. Use **--vlm-conf=/home/vlc/vlc.streaming.conf** to point VLC to open - at start - a special file with multiple files description.

#### Special multiple files description configuration file

-   **vlc.streaming.conf**

Using this config file we try to cast 2 video files: **2007.avi** and **Jumper.avi**. To do this, we must describe 2 channels: *channel1* and *channel2*, set the input, and set the output format (we try to multicast this):

      new channel1 broadcast enabled
      setup channel1 input file:////home/vlc/2007.avi loop
      setup channel1 output #rtp{access=udp,mux=ts,dst=224.255.1.1,port=1234,sdp=sap,sap,group="Video",name="2007 Movie"}

      new channel2 broadcast enabled
      setup channel2 input file:////home/vlc/Jumper.avi loop
      setup channel2 output #rtp{access=udp,mux=ts,dst=224.255.1.2,port=1234,sdp=sap,sap,group="Video",name="Jumper Movie"}

      control channel1 play
      control channel2 play

### Command Line Examples {#streaming-howto-command-line-examples}

Examples for advanced use of VLC's stream output (transcoding, multiple streaming, etc...)

#### Transcoding

Transcode a stream to Ogg Vorbis with 2 channels at 128kbps and 44100Hz and save it as *foobar.ogg*:

    % vlc -I dummy -vvv input_stream --sout
    "#transcode{vcodec=none,acodec=vorb,ab=128,channels=2,samplerate=44100}:file{dst=foobar.ogg}"

Transcode the input stream and send it to a multicast IP address with the associated SAP announce:

    % vlc -vvv input_stream --sout
    '#transcode{vcodec=mp4v,acodec=mpga,vb=800,ab=128,deinterlace}:
    rtp{mux=ts,dst=239.255.12.42,sdp=sap,name="TestStream"}'

Display the input stream, transcode it and send it to a multicast IP address with the associated SAP announce:

    % vlc -vvv input_stream --sout
    #duplicate{dst=display,dst="transcode{vcodec=mp4v,acodec=mpga,vb=800,
    ab=128,deinterlace}:rtp{mux=ts,dst=239.255.12.42,sdp=sap,name="TestStream"}"}'

Transcode the input stream, display the transcoded stream and send it to a multicast IP address with the associated SAP announce:

    % vlc -vvv input_stream --sout
    '#transcode{vcodec=mp4v,acodec=mpga,vb=800,ab=128,deinterlace}:
    duplicate{dst=display,dst=rtp{mux=ts,dst=239.255.12.42,sdp=sap,name="TestStream"}}'

To receive the input stream that is being multicasted above on a client:

    % vlc rtp://239.255.12.42

##### More complex transcoding example

Stream a SDI card to H.264 and AAC in TS on UDP

    % cvlc -vvv --live-caching 2000 decklink://
    --decklink-audio-connection embedded --decklink-aspect-ratio 16:9 --decklink-mode hp50
    --sout-x264-preset slow --sout-x264-tune film --sout-transcode-threads 8 --no-sout-x264-interlaced
    --sout-x264-keyint 50 --sout-x264-lookahead 100 --sout-x264-vbv-maxrate 6000 --sout-x264-vbv-bufsize 6000
    --sout '#transcode{vcodec=h264,vb=6000,acodec=mp4a,aenc=fdkaac,ab=256}:std{access=udp,mux=ts,dst=192.168.2.1:1234}'

#### Multiple streaming

Send a stream to a multicast IP address and a unicast IP address:

    % vlc -vvv input_stream
    --sout '#duplicate{dst=rtp{mux=ts,dst=239.255.12.42,sdp=sap,name="TestStream"},dst=rtp{mux=ts,dst=192.168.1.2}}'

Display the stream and send it to two unicast IP addresses:

    % vlc -vvv input_stream
    --sout '#duplicate{dst=display,dst=rtp{mux=ts,dst=192.168.1.12},dst=rtp{mux=ts,dst=192.168.1.42}}'

Send parts of a multiple program input stream:

    % vlc -vvv multiple_program_input_stream
    --sout'#duplicate{dst=rtp{mux=ts,dst=239.255.12.42},select="program=12345",dst=rtp{mux=ts,dst=239.255.12.43},select="video,program=1234-2345"}'

This command sends the program of the input stream which id is 12345 to 239.255.12.42 and all video programs with id between 1234 and 2345 to 239.255.12.43.

#### Transcoding and multiple streaming

Transcode the input stream, display the transcoded stream and send it to a multicast IP address with the associated SAP announce and an unicast IP address:

    % vlc -vvv input_stream --sout
    '#transcode{vcodec=mp4v,acodec=mpga,vb=800,ab=128,deinterlace}:
    duplicate{dst=display,dst=rtp{mux=ts,dst=239.255.12.42,sdp=sap,name="TestStream"},
    dst=rtp{mux=ts,dst=192.168.1.2}}'

Display the input stream, transcode it and send it to two unicast IP addresses:

    % vlc -vvv input_stream --sout  '#duplicate{dst=display,dst="transcode{vcodec=mp4v,acodec=mpga,vb=800,ab=128}:
    duplicate{dst=rtp{mux=ts,dst=192.168.1.2},dst=rtp{mux=ts,dst=192.168.1.12}"}'

Send the input stream to a multicast IP address and the transcoded stream to another multicast IP address with the associated SAP announces:

    % vlc -vvv input_stream --sout
    '#duplicate{dst=rtp{mux=ts,dst=239.255.1.2,sdp=sap,name="OriginalStream"},
    dst="transcode{vcodec=mp4v,acodec=mpga,vb=800,ab=128}:
    rtp{mux=ts,dst=239.255.1.3,sdp=sap,name="TranscodedStream"}"}'

##### More complex multi-transcoding example

Take a SDI input, and transcode it twice, once in HD, and one in SD and send both on udp.

    % cvlc -vv --live-caching 2000
    --decklink-audio-connection embedded --decklink-aspect-ratio 16:9 --decklink-mode hp50 decklink://
    --sout-x264-preset fast --sout-x264-tune film --sout-transcode-threads 24 --no-sout-x264-interlaced
    --sout-x264-keyint 50 --sout-x264-lookahead 100 --sout-x264-vbv-maxrate 4000 --sout-x264-vbv-bufsize 4000
    --sout '#duplicate{dst="transcode{vcodec=h264,vb=6000,acodec=mp4a,aenc=fdkaac,ab=256}:std{access=udp,mux=ts,dst=192.168.1.2:4013}",
    dst="transcode{height=576,vcodec=h264,vb=2000,acodec=mp4a,aenc=fdkaac,ab=128}:std{access=udp,mux=ts,dst=192.168.1.2:4014}"}'

Take a SDI input, and restreaming it once in raw and transcoding it for the second

    % cvlc -vv --live-caching 2000
    --decklink-audio-connection embedded --decklink-aspect-ratio 16:9 --decklink-mode hp50 decklink://
    --sout-x264-preset fast --sout-x264-tune film --sout-transcode-threads 24 --no-sout-x264-interlaced
    --sout-x264-keyint 50 --sout-x264-lookahead 100 --sout-x264-vbv-maxrate 4000 --sout-x264-vbv-bufsize 4000
    --sout '#duplicate{dst="transcode{vcodec=h264,vb=6000,acodec=mp4a,aenc=fdkaac,ab=256}:std{access=udp,mux=ts,dst=192.168.1.2:4013}",
    dst="std{access=udp,mux=ts,dst=192.168.1.2:4014}"}'

#### HTTP streaming

Stream in HTTP:

-   on the server, run:

&nbsp;

    % vlc -vvv input_stream --sout '#standard{access=http,mux=ogg,dst=server.example.org:8080}'

-   on the client(s), run:

&nbsp;

    % vlc 0

Transcode and stream in HTTP:

    % vlc -vvv input_stream --sout '#transcode{vcodec=mp4v,acodec=mpga,vb=800,ab=128}:
    standard{access=http,mux=ogg,dst=server.example.org:8080}'

Recording a live video stream:

    % vlc 0 --sout="#duplicate{dst=std{access=file,mux=asf,
    dst='C:\test\test.asf'},dst=nodisplay}"

For example, if you want to stream an audio CD in Ogg/Vorbis over HTTP:

    % vlc -vvv cdda:/dev/cdrom
    --sout '#transcode{acodec=vorb,ab=128}:standard{access=http,mux=ogg,dst=server.example.org:8080}'

#### RTSP live streaming

Stream with RTSP and RTP:

-   Run on the server:

&nbsp;

    % vlc -vvv input_stream --sout '#rtp{dst=192.168.0.12,port=1234,sdp=rtsp://server.example.org:8080/test.sdp}'

-   Run on the client(s):

&nbsp;

    % vlc rtsp://server.example.org:8080/test.sdp

#### RTSP on-demand streaming

See Documentation:Streaming HowTo/VLM.

#### MMS / MMSH streaming to Windows Media Player

    % vlc -vvv input_stream --sout '#transcode{vcodec=DIV3,vb=256,scale=1,acodec=mp3,ab=32,
    channels=2}:std{access=mmsh,mux=asfh,dst=:8080}'

VLC media player can connect to this by using the following url: **mmsh://server_ip_address:8080**. Windows Media Player can connect to this by using the following url: **mms://server_ip_address:8080**.

#### Use the *es* module

See also: ES

Separate audio and video in two PS files:

    % vlc -vvv input_stream --sout '#es{access=file,mux=ps,url_audio=audio-%c.%m,url_video=video-%c.%m}'

Extract the audio track of the input stream to a TS file:

    % vlc -vvv input_stream --sout '#es{access_audio=file,mux_audio=ts,url_audio=audio-%c.%m}'

Stream in unicast the audio track on a port and the video track on another port (NOTE: This will not only work with VLC 0.8.6 or older - FIXME?):^\[Please\ check\ this\]^

-   on the server side:

&nbsp;

    % vlc -vvv input_stream --sout '#es{access=rtp,mux=ts,url_audio=192.168.1.2:1212,
    url_video=192.168.1.2:1213}'

-   on the client side:
    -   to receive the audio:

&nbsp;

    % vlc udp://@:1212

-   -   to receive the video:

&nbsp;

    % vlc udp://@:1213

Stream in multicast the video and dump the audio in a file:

    % vlc -vvv input_stream --sout '#es{access-video=udp,mux-video=ts,dst-video=239.255.12.42,
    access-audio=file,mux-audio=ps,dst-audio=audio-%c.%m}'

Note: You can also combine the *es* module with the other modules to set-up even more complex solution.

#### Keeping the stream open

    % vlc -vvv input_stream -sout-keep
    -sout=#transcode{acodec=mp3}:duplicate{dst=display{delay=6000},
    dst=gather:std{mux=mpeg1,dst=:8080/stream.mp3,access=http},select="novideo"}

The basic transcoding is an mp3 stream from the file you select (if it is a video file, then the video is ignored). It is streamed via http to localhost:8080/stream.mp3

The combination of :sout-keep and dst=gather:std mean that the stream is kept open and subsequent items are played through the same stream.

#### Using VLC as a reflector

Taking a UDP input and resending it once raw via IPv6 multicast, and once in HLS

    % cvlc -vvv udp://@:4013 --ttl 60
    --sout '#duplicate{dst=std{access=http,mux=ts,dst=[::]:3013}",
    dst=std{access=udp,mux=ts,dst=ffe2::1]:2013},
    dst=std{access=livehttp{seglen=5,delsegs=true,numsegs=5,index=/path/to/stream.m3u8,
    index-url=0,mux=ts{use-key-frames},dst=/path/to/stream-########.ts}}}

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Easy Streaming {#streaming-howto-easy-streaming}

*This page is outdated. Please see Documentation:Streaming HowTo New#Streaming using the GUI for updated streaming tutorials.*

#### Intro

The easier way to start streaming with VLC is by using one of the graphical user interfaces. These are the wxWindows and skinnable interfaces for Windows and GNU/Linux and the Mac OS X native interface.

#### Streaming using the Wizard

The *Streaming/Transcoding Wizard* leads you step by step through the process of streaming your media on a network or saving it to your hard drive. This *Wizard* offers easy to use menus but provides a restricted set of options.

Note: The wizard is only available on the wxWindows interface.

##### Launching the wizard

To launch the *Streaming/Transcoding Wizard* open the "File" menu, and select the Wizard menu item.

 Launching the Wizard

##### Wizard dialog

First select the type of task:

-   *Stream to network*: Choose this option if you want to stream media on a network.
-   *Transcode/Save to file*: Choose this option if you want to change a file's audio codec and/or video codec, its bitrate, and/or encapsulation method.

 The Wizard Dialog

##### Input selection

Select a stream (such as a file, a network stream, a disk, a capture device ...) by selecting the *Choose...* dialog or an existing item in your playlist, using the *Existing playlist item* option.

*Partial Extract*: To read only part of the stream, check the "Enable" checkbox and choose a start and end date (in seconds). This option should only be used with streams you can control such as files or discs but not network streams or capture devices.

 Wizard input selection

 Wizard input selection from playlist

##### Streaming methods

If you chose *Stream to network* option, you can now specify the streaming method. Available methods are:

-   *RTP/UDP Unicast*: Stream to a single computer. Enter the client's IP address (in the 0.0.0.0 - 223.255.255.255 range).
-   *RTP/UDP Multicast*: Stream to multiple computers using multicast. Enter the IP address of the multicast group (in the 224.0.0.0 to 239.255.255.255 range).
-   *HTTP*: Stream by using the HTTP protocol. If you leave the *Destination* text box empty, VLC will listen on all the network interfaces of the server on port 8080. Specify an address, port and path on which to listen using the following syntax \[ip\]\[:port\]\[/path\]. For instance, *192.168.0.1:80/stream* will make VLC listen on the interface carrying the 192.168.0.1 IP address, on the 80 TCP port, in the /stream *virtual file*.

 Wizard streaming method

##### Transcoding options

If you chose the *Transcode/Save to file* option, you can now specify the new audio and video codecs and bitrates you want you input converted to.

(See \)

 Wizard transcode

##### Encapsulation method

Choose the method format. The UDP streaming methods require MPEG TS encapsulation. The HTTP streaming method can be used with the MPEG PS, MPEG TS, MPEG 1, OGG, RAW or ASF encapsulation. Saving to a file can be done using any encapsulation format compatible with the chosen codecs.

(See \)

 Wizard encapsulation method

##### Streaming options

If you chose to *Stream to network* you can now specify several options.

-   *Time To Live (TTL)* This sets the numbers of routers your stream can go through, for UDP unicast and unicast access methods. If you do not know what this means, you should leave the default value. Note: With UDP multicast, the default TTL is set to 1, meaning that your stream won't get across any router. You may want to increase it if you want to route your multicast stream.
-   *SAP Announce* To advertise your stream over the network when using the UDP streaming method, using the SAP protocol, enter the name of the stream in the text input and check the checkbox. This is NOT available for the HTTP streaming method.

 Wizard streaming options

##### Save to file destination

If you chose *Transcode/Save to file* you can now specify the file you want to save the stream to.

 Wizard save file - wxWindows interface

You can now select the *Finish* button to start streaming/converting the source.

#### Streaming using the GUI

##### Introduction

A second way to set up a streaming instance using VLC is using *Stream Output* panel in the *Open...* dialog of the wxWindows (Windows / GNU Linux), skinnable (Windows / GNU Linux) and MacOS X interfaces. Streaming methods and options used 99% of time should be available in this panel.

To stream the opened media, check the "Stream output" or "Stream/Save" checkbox in the "Open File/Disc/Network Stream/Capture Device" dialog and click on the "Settings" button.

 Open file dialog - wxWindows interface

 Open file dialog - Mac OS X interface

Note that "Capture" is not available as an option in Mac OSX because VLC does not support live streaming of audio or video under Mac OSX.

##### The Stream Output dialog

 Stream output dialog - wxWindows interface

 Stream output dialog - wxWindows interface

###### Stream Output MRL

On the wxWindows interface, a text box displays the *Stream Output MRL* (Media Resource Locator). This is updated as you change options in the Stream output dialog. For more information on how to edit the *Stream Output MRL* read \.

###### Output methods

-   *Play locally*: display the stream on your screen. This allows you to display the stream you are actually streaming. Effects of transcoding, rescaling, etc. can be monitored locally using this function.
-   *File*: Save the stream to a file. The *Dump raw input* option allows you to save the input stream as it is read by VLC, without any processing.
-   *HTTP*: Use the HTTP streaming method. Specify the IP address and TCP port number on which to listen.
-   *MMSH*: This access method allows you to stream to Microsoft Windows Media Player. Specify the IP address and TCP port number on which to listen. Note: This will only work with the *ASF* encapsulation method.
-   *UDP*: Stream in unicast by providing an address in the 0.0.0.0 - 223.255.255.255 range or in multicast by providing an address in the 224.0.0.0 - 239.255.255.255 range. It is also possible to stream to IPv6 addresses. Note: This will only work with the *TS* encapsulation method.
-   *RTP*: Use the Real-Time Transfer Protocol. Like UDP, it can use both unicast and multicast addresses.

Note: UDP, HTTP, MMSH, and RTP methods require you to select the *Stream* option on the MacOS X interface.

(See \)

###### Encapsulation method

Select an encapsulation method that fits the codecs and access method of your stream, among MPEG TS, MPEG PS, MPEG 1, OGG, Raw, ASF, AVI, MP4 and MOV. (See \)

###### Transcoding options

Enable video transcoding by checking the "Video Codec" checkbox. Choose a codec from the list. You can also specify an average bitrate and scale the input. (See \)

Enable audio transcoding by checking the "Audio Codec" checkbox. Choose a codec from the list. You can also specify an average bitrate and the number of audio channels to encode. (See \)

###### Miscellaneous options

Select methods to announce your stream. You can use SAP (Service Announce Protocol) or SLP (Service Location Protocol). You must also specify a channel name. The Mac OS X interface also allows you to export the description (SDP) file of a RTP session using the internal HTTP or RTSP server of VLC, or as a file. This can be done using the according checkboxes. The *SDP URL* text box allows to give the url or destination where the SDP file will be available.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Easy Streaming Newer Versions {#streaming-howto-easy-streaming-newer-versions}

#### Introduction

Since the documentation on streaming is fairly old, this wiki page was written to show how to do streaming on newer versions of VLC Media Player. EDIT: This page looked incomplete, and I figured out a way that worked for me on my particular system, so I thought I'd document it in the hopes that it helps someone else in the future. My setup is as follows:

-   Server: Windows Vista Machine running VLC 1.1.11 (IP Address: 192.168.2.2)
-   Client: Windows 7 Machine running VLC 1.1.11 (IP Address: 192.168.2.4)
-   Two computers are on the same subnet (192.168.2.X). I am able to ping from each machine to the other.

Goal: I have a bunch of video files ripped from DVDs that I want to share between my server and my client(s). This is simply to be able to keep all my movies in one central location.

#### SERVER SETUP: Streaming using the streaming dialog

-   Launch VLC and then push Media -\> Streaming... (You should see a dialog like the one below)

-   On the window that pops up click on the tab of the media you want to stream from. I chose to stream from a file (a .vob file). I add this file to the "File Selection" list.
    -   I left "Use a subtitles file" unchecked
-   Hit the "Stream" button at the bottom of the dialog. This pops up a new streaming options dialog. The streaming options dialog has 3 sections: Source, Destinations, and Options.

##### Source Dialog

 Source is already filled in with the file that I chose already, so I hit the Next button

##### Destinations

-   Under Destinations, I selected "RTP / MPEG Transport Stream", since this is an awesome way to send data across the network (that's what RTP is all about).
-   In my case, I want to see the movie both on the server and the client (it gives me warm fuzzies to see the movie in both places), so I check the "Display locally" checkbox
-   Under "Transcoding options", I uncheck "Activate Transcoding" because I happen to know that my videos are already encoded just fine (and when I tried transcoding, that didn't seem to work so well for me)

-   Now click "Add". Specify the IP address of the client (in my case 192.168.2.4 - if you don't know what yours is, open a command prompt (cmd.exe), and run "ipconfig /all")

-   Hit "Next" to go to the final options screen

##### Options

-   Under "Options", I selected "Stream all elementary streams" (not completely necessary it turns out...this probably sends more than I really have to), and I also checked "SAP announce" and gave it a name (I chose the name of the video file...seemed logical) and a group name (doesn't seem to be all that important)

-   Hit "Stream", and your movie should start playing locally (and it should start streaming)

#### CLIENT SETUP: Receiving the stream

Since I configured SAP in the server, it's easy to open the stream on the client: I just open up the media browser view of VLC (by clicking on the button next to the full screen button) and look under "Local network". I see the name of my SAP stream show up, and I double click it. Voila!! Streaming video! \*\*Football style chest bump\*\*

Note: you can start and stop the stream on the client, just as long as you don't catch up to the server. Pretty nice!

### Receive and Save a Stream {#streaming-howto-receive-and-save-a-stream}

#### Receive a stream with VLC

##### Receive an unicast stream

    % vlc -vvv rtp://

##### Receive a multicast stream

    % vlc -vvv rtp://@239.255.12.42

where **239.255.12.42** is the multicast IP address you want to join.

##### Receive an HTTP/FTP/MMS stream

Use one of the following command lines:

    % vlc -vvv 0

where **0** is the HTTP address of the stream;

    % vlc -vvv ftp://example/stream.xyz

where **[ftp://example/stream.xyz](ftp://example/stream.xyz)** is the FTP address of the stream;

    % vlc -vvv mms://viptvr.yacast.fr/encoderfranceinfo

where **[mms://viptvr.yacast.fr/encoderfranceinfo](mms://viptvr.yacast.fr/encoderfranceinfo)** is the MMS address of the stream.

##### Receive a RTP stream available through RTSP

    % vlc -vvv rtsp://www.hardradio.com/tonbeme.mov

where **rtsp://www.hardradio.com/tonbeme.mov** is the address of the stream.

##### Receive a stream described by an SDP file

    % vlc -vvv 0

#### Save a stream with VLC

VLC can save the stream to the disk. In order to do this, use the Stream Output of VLC: you can do it via the graphical interface (Media \[menu\] → streaming) or use the [record button](http://www.howtogeek.com/howto/2686/how-to-copy-a-dvd-with-vlc-1.0/), or you can add to the command line the following argument:

    --sout file/muxer:stream.xyz

where:

-   **muxer** is one of the formats supported by VLC's stream output, i.e. :
    -   **ogg** for OGG format,
    -   **ps** MPEG2-PS format,
    -   **ts** for MPEG2-TS format.
-   and **stream.xyz** is the name of the file you want to save the stream to, with the right extension.

For example:

    vlc your_input_file_or_stream_here --sout=file/ps:go.mpg

This is short hand for the more verbose

    vlc your_input_file_or_stream_here --sout="#std{access=file,mux=ps,dst=go.mpg}"

NB that you must choose a muxer that supports your stream type. See Transcode#Compatibility_issues

It can also be quite helpful to look at the settings VLC uses when it records using its record button. For example, in the logs you might see something like this:

...: Using record output \`std{access=file,mux='ps',dst='C:\\vlc-record-2010\_\_E-.mpg'}'

Which gives you a hint/clue as to how to record your current stream. In this case this would translate into --sout "#std{access=file... on the command line.

#### Receive a stream with a set-top-box

Some set-top-boxes with Ethernet cards can receive MPEG2-TS streams over UDP and support multicast.

Set-top-boxes known to work with VLC are:

-   [Pace](http://www.pace.co.uk) set top boxes. (Pace Micro DSL 4000)
-   [Aminocom](http://www.aminocom.com) set top boxes. (all the models with mpeg2)
-   tuxia / gct-allwell (mpeg4 and mpeg2) sigma designs8174 chipset
-   i3micro mood200 (mpeg4 and mpeg2 in transport streams)
-   ps3 media server streams using VLC (or mencoder) to the PS3

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Stream a DVB Channel {#streaming-howto-stream-a-dvb-channel}

Note: This is possible under GNU/Linux only.

#### Install the DVB drivers

If you want to be able to stream from a DVB card (a satellite card or a digital terrestial TV card), you need to install the DVB drivers:

-   if you use a Linux 2.6.x kernel, you just need to select the right modules in your kernel configuration.
-   if you use a dvb-card Technisat SkyStar2 rev. 2.8, you must download the latest release of the DVB drivers from the [DVB-S(S2) drivers for Linux](http://mercurial.intuxication.org/hg/s2-liplianin/).
-   if you are using a Linux 2.4.x kernel, you must download the latest release of the DVB drivers from the [DVB drivers download page](http://www.linuxtv.org/download/dvb/) of the [LinuxTV](http://www.linuxtv.org/) Project.

The following sections assume that you have a working linux-dvb installation, either from stock kernel 2.6 or from kernel 2.4 with DVB patches. If you have any problem with the linux-dvb drivers, please report the problem to the maintainers of the drivers, not to us. Thanks.

#### Stream with VLS

Note: VLS is currently deprecated and hasn't been maintained for years. It is strongly advised to use VLC instead, which now supports the same features as VLS, and many more. The only advantage of VLS is to support the dvbrc file syntax, and it requires a bit less CPU horsepower. However, we do not support VLS any longer.

Put a **.dvbrc** file containing the DVB channels (satellite or digital terrestial TV channels) you want to stream in your home directory (some are provided in the *libdvb* tarball for the satellite channels).

Run VLS with the following command line:

    % vls -vv -d udp:192.168.0.42 dvb:"EUROSPORT" --ttl 12

where:

-   **"EUROSPORT"** is the channel you want to stream as written in your **\~/.dvbrc** file,
-   **192.168.0.42** is either:
    -   the IP address of the machine you want to unicast to;
    -   or the DNS name the machine you want to unicast to;
    -   or a multicast IP address.
-   **12** is the value of the TTL (Time To Live) of your IP packets (which means that the stream will be able to cross 11 routers).

#### Stream with VLC

Note: VLC has many more features than VLS. First you can use the advanced stream output options such as transcoding and all kinds of output supports. Second VLC can take advantage of the Common Interface supported by some DVB adapters to descramble one or several services. Currently released versions of VLC only support the low-level API so some adapters won't work (budget-ci cards work, twinhan doesn't). Some CAM modules aren't compatible with some DVB cards, check the linux-dvb documentation for more information. So-called "professional" CAM modules are able to descramble up to twelve services, whereas customer-oriented modules are often limited to one or two services unless otherwise specified.

VLC must be compiled with --enable-dvb and you need the linux-dvb headers installed in your system. An example command-line is as follows:

    % vlc -vvv --color --ttl 12  --ts-es-id-pid --programs=8508,8505 dvb:
      --dvb-frequency=11739000 --dvb-srate=27500000 --dvb-voltage=13
      --sout-standard-access=udp --sout-standard-mux=ts --sout
      '#duplicate{dst=std{dst=address1},select="program=8508",dst=std{dst=address2},select="program=8505"}'

The example above shows the minimum set of options needed to stream out two services. Here is a list of frontend options, depending on the frontend type:

-   *common options*
    -   **dvb-adapter**: specifies the adapter to use in case you have several adapters in your machine (by default use adapter 0)
    -   **dvb-device**: specifies the name of the DVB device to use (should not be needed with a standard linux-dvb installation)
    -   **dvb-srate**: specifies the symbol rate of the modulated signal, in symbols/s
    -   **dvb-inversion**: specifies whether the signal is inverted or not (default is automatic detection)
    -   **dvb-budget-mode**: enters a special mode where all PIDs are retrieved by the driver; it should no longer be necessary as VLC should filter wanted PIDs
-   *satellite frontend (QPSK)*
    -   **dvb-frequency**: specifies the frequency to tune to in kHz; according to the frequency range, VLC auto-detects the band to use: S (2.5-2.7 GHz), C-lower (3.4-4.2 GHz), C-higher (4.5-4.8 GHz), Ku (10.7-13.25 GHz) or direct BIS frequency (0.95-2.15 GHz); it is mandatory to supply the **dvb-srate** option to satellite frontends
    -   **dvb-voltage**: specifies the voltage to apply on the IF; most LNBs behave differently when supplied with 13 V or 18 V; universal LNBs select vertical polarity with 13 V and horizontal with 18 V; you can also select 0 V if your LNB has another power supply (default is 13 V)
    -   **dvb-tone**: specifies whether to send a 22 kHz pulse tone to the LNB; universal LNBs switch to high-band when this pulse is sent; by default VLC automatically adopts the correct behaviour if the frequency supplied is in the Ku band (other bands do not need this)
    -   **dvb-fec**: specifies the code-rate to use for Forward Error Correction; type in the first number of the code-rate, for 2/3 use --dvb-rate=2, etc. (default is 9, meaning automatic detection)
    -   **dvb-high-voltage**: enables a special mode of the DVB adapter to compensate for the voltage loss in very long cables (AFAIK it is present in the API, but no DVB adapter actually implements it)
    -   **dvb-lnb-lof1, dvb-lnb-lof2, dvb-lnb-slof**: specifies the frequencies of the first and second local oscillators, and the frequency at which the 22 kHz pulse should be activated to enable the second oscillator; by default VLC uses the values for universal LNBs if the frequency supplied is in the Ku band (other bands do not need this)
-   *cable frontend (QAM)*
    -   **dvb-frequency**: specifies the frequency to tune to in Hz; it is mandatory to supply the **dvb-srate** option to cable frontends
    -   **dvb-modulation**: specifies the modulation of the analog signal; valid values are -1 (QPSK), 0 (automatic QAM, default), 16 (QAM16), 32 (QAM32), 64 (QAM64) 128 (QAM128), 256 (QAM256)
-   *terrestrial frontend (OFDM)*
    -   **dvb-frequency**: specifies the frequency to tune to in Hz; it is mandatory to supply the **dvb-bandwidth** option, all other parameters are optional
    -   **dvb-bandwidth**: specifies the bandwidth of the OFDM channel (6, 7 or 8 MHz depending on the country)
    -   **dvb-hierarchy**: specifies if the OFDM channel uses hierarchic information; allowed values are -1 (no hierarchy), 0 (automatic, default), 1, 2 and 4
    -   **dvb-code-rate-hp, dvb-code-rate-lp**: specifies the code-rate to use for higher and lower hierarchies respectively (default auto, same syntax as **dvb-fec**)
    -   **dvb-guard**: specifies the guard interval; valid values are 0 (automatic, default), 4 (1/4), 8 (1/8), 16 (1/16) and 32 (1/32)
    -   **dvb-transmission**: specifies the transmission mode; valid values are 0 (automatic, default), 2 (2K) and 8 (8K)

We also ought to explain the other non-dvb-specific options of the example command-line:

-   **ts-es-id-pid**: this option is necessary if you use the **#duplicate** stream output filter to split the multiplex in several outputs; there is no need to use **#duplicate** neither **ts-es-id-pid** if you have one program only
-   **programs, program, sout-all**: there are several ways of specifying the services to select (and optionally descramble):
    -   **programs**: used to specify one or serveral programs to select; VLC selects all known elementary streams of these programs; this is the currently recommended way
    -   **program**: used to specify one program to select; it differs from using **programs** with only one program in that this option only select the first audio stream, and no subtitle stream; it should be used if you plan to switch programs and audio with a GUI
    -   **sout-all**: tells VLC to select all programs; this is discouraged because of the extra CPU load needed to demultiplex unwanted programs, and because it is not compatible with CAM descrambling
-   The other options are standard stream output options and are described in the other chapters of this documentation.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Stream a DVD {#streaming-howto-stream-a-dvd}

Note: Under Unix/Linux, you must have write access to the device corresponding to your DVD drive. For that, you should be in the *disk* or *cdrom* group (look at the permissions in **/dev**). If you're not, add yourself to the group:

    # adduser your_login disk_or_cdrom

and then restart your session.

#### Stream a DVD with VLC

    % vlc -vvv --color dvdsimple:/dev/dvd --sout udp://192.168.0.12 --ttl 12 --sout-all

where:

-   **/dev/dvd** is the name of your DVD drive (put **D:** under Windows if **D** is the letter of your DVD drive) or the directory where you copied your DVD,
-   **192.168.0.42** is either:
    -   the IP address of the machine you want to unicast to;
    -   or the DNS name the machine you want to unicast to;
    -   or a multicast IP address.
-   **12** is the value of the TTL (Time To Live) of your IP packets (which means that the stream will be able to cross 11 routers).
-   **sout-all** allows you to stream all soundtracks and subtitles

If you want to stream the DVD continuously, add the **--loop** option.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Stream a File {#streaming-howto-stream-a-file}

#### Stream a file with VLC

    % vlc -vvv video1.xyz --sout udp:192.168.0.42 --ttl 12

where:

-   **video1.xyz** is the file you want to stream,
-   **192.168.0.42** is either:
    -   the IP address of the machine you want to unicast to;
    -   or the DNS name the machine you want to unicast to;
    -   or a multicast IP address.
-   **12** is the value of the TTL (Time To Live) of your IP packets (which means that the stream will be able to cross 11 routers).

If you want to stream the file continuously, add the **--loop** option.

Of course, you can add more options (like transcoding, or streaming to a TCP port, etc.), but this should get you started.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Stream from a DV Camcorder {#streaming-howto-stream-from-a-dv-camcorder}

Note: This is possible under GNU/Linux only.

#### Install the libraw1394 and libavc1394

If you want to be able to stream from a DV camcorder, then you need to install the libraries libraw1394 and libavc1394:

-   if you use a Fedora Core distribution then you just need to install the libraries using:

&nbsp;

    % yum update
    % yum install libraw1394 libavc1394'

-   if you want to install the libraries from the source then you must download them from the [libraw1394](http://www.linux1394.org/) and [libavc1394](http://sourceforge.net/projects/libavc1394) from their projects website.

&nbsp;

-   if you have a distribution that uses [udev](http://kernel.org/pub/linux/utils/kernel/hotplug), then you must add/change the following line to the file 50-udev.rules in your /etc/udev/rules.d directory.

&nbsp;

    % vi /etc/udev/rules.d/50-udev.rules
    # IEEE1394 (firewire) devices (must be before raw devices below)
    KERNEL=="raw1394",              NAME="%k"
    KERNEL=="dv1394",               NAME="dv1394/%k"
    KERNEL=="video1394*",           NAME="video1394/%n"

The following sections assume that you have a working linux installation with the IEEE 1394 (Firewire) libraries installed, either manually from the source code or through your distributions upgrade mechanism.

#### Stream with DV

Connect the DV camcorder with a Firewire cable to your computer, and check the creation of the file **/dev/raw1394**.

Run VLC with the following in one command line:

    % vlc -vvv dv/rawdv:///dev/raw1394 --dv-caching 10000 --sout
    '#transcode{vcodec=WMV2,vb=512,scale=1,acodec=mp3,ab=192,channels=2,fps=25.0}:
    std{access=mmsh,mux=asfh,url=:8080}'

where:

-   **dv/rawdv://** is the DV input and **/dev/raw1394** the device file,
-   **dv-caching** is the delay is milliseconds (ms) (start with a high value, 10s or so, and lower it later),
-   **sout** is the stream output chain that is used to stream the DV camcorder as a multimedia stream over the network. The **transcode** syntax is explained in the chapter about transcoding. The example as given above generates a multimedia stream that is compatible with Windows Media Player,
-   **sout-transcode-fps** is the number of pictures per second **25.0** that the transcode module should generate of the requested audio/video codec.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Stream from Encoding Cards and Other Capture Devices {#streaming-howto-stream-from-encoding-cards-and-other-capture-devices}

#### Hardware encoding cards

Note: This is possible under GNU/Linux only.

VideoLAN supports two kinds of MPEG-2 encoding cards:

-   Hauppauge WinTV-PVR-250 and WinTV-PVR-350,
-   Visiontech Kfir.

The Hauppauge WinTV-PVR-250/350 gives much better results and is cheaper than the Visiontech Kfir.

##### Stream with the Hauppauge WinTV-PVR-250/350 card

###### Install the drivers

First, you will have to patch your kernel (version 2.4) to support the v4l2 API (Video 4 Linux version 2). The patch is available on the [Video4Linux HQ](http://bytesex.org/v4l/). If you use a 2.6 kernel, you only need to build I2C support and the BT848 Video For Linux module.

Once your kernel is ready, install the CK version (currently in development) of the Linux drivers for the Hauppauge WinTV-PVR-250/350. They are hosted on [ivtv ck](http://67.18.1.101/~ckennedy/ivtv). You will need to patch your kernel to use it with a 2.4. You can also use the CVS version available here: [ivtv.sourceforge.net](http://ivtv.sourceforge.net/) (this version is not developped anymore). Then, you will have to create the device and load the modules; for this, please refer to the documentation shipped with the drivers.

###### Stream with VLC

Note: You must add **--enable-pvr** to **./configure** to use this feature.

    % vlc -vvv --color pvr:///dev/video0:norm=secam:size=720x576:frequency=576250:bitrate=3000000:maxbitrate=4000000 --cr-average 1000 --sout '#rtp{mux=ts,dst=192.168.0.42,port=5004}' --ttl 12

where:

-   **/dev/video0** is the device corresponding to the encoding card,
-   **norm=secam** is name of the standard of the analogic signal (possible values are pal, secam, and ntsc),
-   **size=720x576** is the size of the video you want to stream,
-   **frequency=567250** is the frequency in kHz of the channel you want to stream,
-   **bitrate=3000000** is the average bitrate of the stream,
-   **maxbitrate=4000000** is the maximum bitrate of the stream,
-   **1000** is a secret value to work around a bug of the card.
-   **192.168.0.42** is either:
    -   the IP address of the machine you want to unicast to;
    -   or the DNS name the machine you want to unicast to;
    -   or a multicast IP address.
-   **12** is the value of the TTL (Time To Live) of your IP packets (which means that the stream will be able to cross 11 routers).

##### Stream with the Visiontech Kfir card

###### Install the drivers

If you want to be able to stream from a Visiontech Kfir card, you need to install its Linux drivers. Download the latest release of the drivers from the [drivers download page](http://www.linuxtv.org/download/mpeg2/) of the \0 LinuxTV web site\].

Uncompress the tarball and follow the instructions written in the *INSTALL* file to compile and install the drivers.

Note: If you have a VIA chipset, you need to disable USB in the BIOS.

###### Stream

    % vlc -vvv --color kfir:///dev/video --sout '#rtp{mux=ts,dst=192.168.0.42,port=5004}' --ttl 12

where:

-   **/dev/video** is the device corresponding to the Kfir card,
-   **192.168.0.42** is either :
    -   the IP address of the machine you want to unicast to;
    -   or the DNS name the machine you want to unicast to;
    -   or a multicast IP address.
-   **12** is the value of the TTL (Time To Live) of your IP packets (which means that the stream will be able to cross 11 routers).

#### Software encoding cards

##### Under GNU/Linux

###### Install the Video for Linux drivers

If you want to stream from an acquisition card or a webcam, a video4linux driver must be available for it. You can find more information about video4linux and supported devices [here](http://www.exploits.org/v4l).

Compile the right module for your device, and insert it into your kernel. Some video4linux modules are shipped with the 2.4.x and 2.6.x Linux kernels, the patch is available on the [Video4Linux HQ](http://bytesex.org/v4l).

You can test your device by using any of the listed programs in the *Video: TV and PVR/DVR* section of [this page](http://www.exploits.org/v4l/).

Note that v4l2 modules will also work with VLC.

###### Stream with VLC

Note: You must add **--enable-v4l** to **./configure** to use this feature.

    % vlc -vvv --color v4l:///dev/video:norm=secam:frequency=543250:size=640x480:channel=0:adev=/dev/dsp:audio=0 --sout '#transcode{vcodec=mp4v,acodec=mpga,vb=3000,ab=256,venc=ffmpeg{keyint=80,hurry-up,vt=800000},deinterlace}:rtp{mux=ts,dst=239.255.12.13,port=5004}' --ttl 12

Note: You can find all transcode options on this page : Advanced Streaming Using the Command Line.

where:

-   **/dev/video** is the device corresponding to your acquisition card or your webcam,
-   **norm=secam** is name of the standard of the analogic signal (possible values are pal, secam, and ntsc),
-   **frequency=543250** is the frequency of the channel in kHz (*Warning:* for VLC \< 0.6.1, Frequency is channel frequency in MHz multiplied by 16),
-   **size=640x480** is the size of the video you want (you can also put the standard size like *subqcif* (128x96), *qsif* (160x120), *qcif* (176x144), *sif* (320x240), *cif* (352x288) or *vga* (640x480)),
-   **channel=0** is the number of the channel (usually 0 is for tuner, 1 for composite and 2 for svideo),
-   **adev=/dev/dsp** is the audio device,
-   **audio=1** is the number of the audio channel (usually 0 is for mono and 1 for stereo),
-   **vcodec=mp4v** is the video format you want to encode in (*mp4v* is MPEG-4, *mpgv* is MPEG-1, and there is also *h263*, *DIV1*, *DIV2*, *DIV3*, *I420*, *I422*, *I444*, *RV24*, *YUY2*),
-   **acodec=mpga** is the audio format you want to encode in (*mpga* is MPEG audio layer 2, *a52* is A52 i.e. AC3 sound),
-   **vb=3000** is the video bitrate in Kbit/s
-   **ab=256** is the audio bitrate in Kbit/s
-   **venc=ffmpeg** allows to set the encoder to use, where:
    -   **keyint=80** is the maximal amount of frames between two key frames
    -   **hurry-up** allows the encoder to decrease the quality of the stream if the CPU can't keep up with the encoding rate
    -   **vt=800000** is the tolerance in kbit/s for the bitrate of the outputted video
-   **deinterlace** tells VLC to deinterlace the video on the fly,
-   **192.168.0.42** is either:
    -   **the IP address of the machine you want to unicast to;**
    -   **or the DNS name the machine you want to unicast to;**
    -   **or a multicast IP address.**
-   **12** is the value of the TTL (Time To Live) of your IP packets (which means that the stream will be able to cross 11 routers).

#### Stream with DirectShow (Windows)

##### Install your peripheral drivers

You need to install your peripherals under Windows with the appropriate drivers. Nothing else is necessary.

##### Stream unicast/multicast with VLC in command line

    % C:\Program Files\VideoLAN\VLC\vlc.exe -I rc --ttl 12 dshow:// vdev="VGA USB Camera" adev="USB Camera" size="640x480" --sout=#rtp{mux=ts,dst=239.255.42.12,port=5004}

Note: You either need to provide the full path to the vlc.exe executable or add its location to the Windows Path variable.

-   **-I rc** is to activate the remote control interface (MS/DOS console)
-   **12** is the value of the TTL (Time To Live) of your IP packets (which means that the stream will be able to cross 11 routers),
-   **vdev="VGA USB Camera"** is the name of the video peripheral that DirectShow will use (this is only an exemple),
-   **adev="USB Camera"** is the name of the audio peripheral,
-   **size="640x480"** is the resolution (you can also put the standard size like *subqcif* (128x96), *qsif* (160x120), *qcif*

(176x144), *sif* (320x240), *cif* (352x288) or *vga* (640x480)).

-   **239.255.42.12** is either:
    -   the IP address of the machine you want to unicast to;
    -   or the DNS name the machine you want to unicast to;
    -   or a multicast IP address.

##### Stream to file(s) with VLC in command line

    % C:\Path\To\vlc.exe -I rc dshow:// :dshow-vdev="Osprey-210 Video Device 1" :dshow-adev="Unbalanced 1 (Osprey-2X0)"  :dshow-caching=200 --sout="#duplicate{dst='transcode{vcodec=h264,vb=1260,fps=24,scale=1,width=640,height=480,acodec=mp4a,ab=96,channels=2,samplerate=44100}:std{access=file,mux=mp4,dst=C:\\Path\\To\\File-1.mp4}',dst='transcode{vcodec=h264,vb=560,fps=24,scale=1,width=427,height=320,acodec=mp4a,ab=96,channels=2,samplerate=44100}:std{access=file,mux=mp4,dst=C:\\Path\\To\\File-2.mp4}'}"

-   **-I rc** is to activate the remote control interface (MS/DOS console)
-   **dshow://...** configures your input capture card / settings
-   **#duplicate{}** multiple output configurations
-   **transcode{}** video/audio codec settings
-   **std{}** output/muxer settings

#### Mac OSX

Note that VLC does not support streaming from live video or audio sources on Mac OSX.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Streaming /  Muxers and Codecs {#streaming-howto-streaming-muxers-and-codecs}

#### Introduction

##### Overview

VideoLAN is a complete software solution for video streaming, developed by students of the [Ecole Centrale Paris](http://www.ecp.fr) and developers from all over the world, under the [GNU General Public License](http://www.gnu.org/copyleft/gpl.html) (GPL). VideoLAN is designed to stream MPEG videos on high-bandwidth networks.

The VideoLAN solution includes:

-   VLS (VideoLAN Server), which can stream MPEG-1, MPEG-2 and MPEG-4 files, DVDs, digital satellite channels, digital terrestrial television channels and live videos on the network in unicast or multicast
-   VLC (initially VideoLAN Client), which can be used as a server to stream MPEG-1, MPEG-2 and MPEG-4 files, DVDs and live videos on the network in unicast or multicast ; or used as a client to receive, decode and display MPEG streams under multiple operating systems

Here is an illustration of the complete VideoLAN solution:

More details about the project can be found on the [VideoLAN Web site](http://www.videolan.org).

##### VideoLAN software

###### VLC Media Player

VLC works on many platforms: Linux, Windows, Mac OS X, BeOS, \*BSD, Solaris, Familiar Linux, Yopy/Linupy and QNX. It can read:

-   MPEG-1, MPEG-2 and MPEG-4 / DivX files from a hard disk, a CD-ROM drive, ...
-   DVDs and VCDs
-   from a satellite card (DVB-S)
-   from a camcorder (DV)
-   MPEG-1, MPEG-2 and MPEG-4 streams from the network sent by VLS or VLC's stream output

VLC can also be used as a server to stream:

-   MPEG-1, MPEG-2 and MPEG-4 / DivX files,
-   DVDs,
-   from an MPEG encoding card,
-   from a camcorder DV,

to:

-   one machine (i.e. to one IP address): this is called *unicast*,
-   a dynamic group of machines that the clients can join or leave (i.e. to a multicast IP address): this is called *multicast*,

in IPv4 or IPv6.

To get the complete list of VLC's possibilities on each platform supported, see the [VLC features page](http://www.videolan.org/vlc/features.html).

Note: VLC doesn't work on Mac OS 9, and probably never will.

###### Mini-SAP-server

You can add a channel information service based on the SAP/SDP standard to the VideoLAN solution. The mini-SAP-server sends announces about the multicast programs on the network in IPv4 or IPv6, and VLCs receive these annouces and automatically add the programs announced to their playlist.

The mini-SAP-server works under Linux and Mac OS X.

#### Muxers and codecs

##### What is a codec ?

To fully understand the VideoLAN solution, you must understand the difference between a *codec* and a *container format*.

A *codec* is a compression algorithm, used to reduce the size of a stream. There are audio codecs and video codecs. MPEG-1, MPEG-2, MPEG-4, Vorbis, DivX, ... are codecs.

##### What is a container format ?

To start off, think of a *container format* as a standard shipping box. You get a box in the mail and you think, "Cool! What's inside?" You don't really care about the box itself, you care about what's in that box. The problem? You can't see into the box. So what do you do? You get a knife and cut it open.

A *container format* follows this same basic idea. It contains one or several streams already encoded by codecs. Very often, there is an audio stream and a video one. AVI, Ogg, MOV, ASF, MP4 ... are container formats. The streams contained can be encoded using different codecs. In a perfect world, you could put any codec in any container format. Unfortunately, there are some incompatibilities. You can find a matrix of possible codecs and container formats on the [features page](http://www.videolan.org/streaming/features.html).

##### Encoding a video

This is the step where you are going to create the shipping box.

1.  Encode your file. This means compressing a file, whether it is audio or video, to another format that normally takes up less physical drive space than the previous format. Common video encoding methods are DivX, MPEG-1, MPEG-2, MPEG-4 ... most common audio encoding method is MP3 or ogg-vorbis.
2.  Mux (or multiplex). This means joining separate parts of the video (or streams) into one file.

##### Playing a video

Now that you have your shipping box, you need to open it before you can see the content. That's exactly what VLC will do. To decode a stream, VLC first *demuxes* it. This means that it reads the container format and separates audio, video, and subtitles, if any. Demuxing files doesn't weaken the video or audio quality, neither does it do anything for them; it simply saves them into separate files, each containing one element of the original file. Then, each of these is passed to a *decoder* that does the mathematical processing to decompress the stream.

There is a particular thing about MPEG:

-   MPEG is a codec. There are several versions of it, called MPEG-1, MPEG-2, MPEG-4, ...
-   MPEG is also a container format, sometimes referred to as MPEG System. There are several types of MPEG: ES, PS, and TS.

For instance, when you play an MPEG video from a DVD, the MPEG stream is actually composed of several streams (called Elementary Streams, ES): there is one stream for video, one for audio, another for subtitles, and so on. These different streams are mixed together into a single Program Stream (PS). So, the .VOB files you can find in a DVD are actually MPEG-PS files. However, this PS format is not adapted for streaming video through a network or by satellite. So, another format called Transport Stream (TS) was designed for streaming MPEG videos through such channels.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Streaming a live video feed to Darwin Streaming Server for Mobile Phones {#streaming-howto-streaming-a-live-video-feed-to-darwin-streaming-server-for-mobile-phones}

#### Streaming a live video to DSS for Mobile Phones with VLC

    % vlc -vvv v4l2:///dev/video0:input=1:width=128:height=96:adev=hw.1,0:samplerate=32000 --sout '#transcode{venc=ffmpeg{keyint=1},vcodec=mp4v,vb=100k,acodec=mp4a,fps=10,ab=8k,channels=1,samplerate=16000}:rtp{mp4a-latm,dst=127.0.0.1,port-audio=20000,port-video=20002,ttl=127,name=CHANNEL,sdp=file:///usr/local/movies/channel.sdp}'

where:

-   **v4l2:///dev/video0** is the video device you want you want to stream,
-   **input=1** is the input channel of the video device (0 - tv tuner, 1 - composite),
-   **width=128:height=96** is the width and height of the input video signal to fetch by VLC,
-   **adev=hw.1,0** is the alsa audio device to capture audio from,
-   **samplerate=32000** is the input sample rate of the audio live feed,
-   **venc=ffmpeg** is the encoder used (in this case it's ffmpeg, but you can use x264),
-   **{keyint=1}** is the advanced ffmpeg encoder switches,
-   **vcodec=mp4v** is video codec used to encode this live video feed (in this case it's MPEG4),
-   **vb=100k** is the video bitrate (100 kbits/s is this case),
-   **acodec=mp4a** is the audio codec used (is this case it's AAC),
-   **fps=10** is the frame rate of the video feed,
-   **ab=8k** is the audio bitrate (is this case 8 kbits/s),
-   **mp4a-latm** is only used for aac audio, it activates a different payload format for aac,
-   **dst=127.0.0.1** is the destination IP, where Darwin Streaming Server is hosted,
-   **ttl=127** is the value of the TTL (Time To Live) of your IP packets (which means that the stream will be able to cross 126 routers),
-   **sdp=file:///usr/local/movies/channel.sdp** is where to create the SDP file for live streaming with Darwin Streaming Server (it should be inside of the DSS movies folder),
-   **name=CHANNEL** is the name of the live video feed.

Tested on Nokia N73 and SE K800.

There is a small problem with some Nokia phones and Darwin Streaming Servers, that need a line to be edited in the created SDP file (for example):

-   **from b=RR:0 to b=RR:800**

After running this command from console, you can access it from your mobile phone or VLC or any player that supports RTSP protocol

-   **rtsp://192.168.2.3/channel.sdp**

where

-   **192.168.2.3** is the IP address of the machine where DSS is running.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Streaming for the iPhone {#streaming-howto-streaming-for-the-iphone}

#### Streaming for the iPhone

This functionality allows you to link VLC's transcoding capability with a segmenter which will in turn create the series of files needed for http live streaming to the iPhone.

Unlike most of VLC's streaming, it doesn't actually stream the files, but assumes that you have your own webserver which will do that job.

Notes from the developer:

-   This has only been tested using H.264 w/MP3 or AAC audio using mux=ts, and raw MP3 using mux=raw
-   I've been mostly an FFMPEG guy till now, so forgive me if my VLC understanding/terminology is somewhat off.
-   This plug-in should support both Live and non-live HTTP Live streaming feeds, depending on the options passed to the module.

Notes from the community:

-   The user running vlc needs to have read/write/delete permissions on the directory where the .m3u8 playlist and .ts file segments will be created and deleted on the webserver.

##### Instructions

The name of the module is livehttp, and is specified by specifying "access=livehttp"

##### Options

splitanywhere= (default: false)

-   Tells livehttp to split the stream anywhere, not just on video keyframes. Currently required to be set to true for audio-only streams and not recommended (probably won't work) for video streams.

seglen= (default: 10)

-   How many seconds of audio/video each segment should contain. Apple recommends 10, I have been using 5.

numsegs=(default: 0)

-   The number of segments to keep in the index file. The default of 0 keeps all segments in the index (which you would want for non-live streaming). For live streaming the specification require at least 3.

delsegs= (default: true)

-   Delete segments as they are no longer needed. If numsegs=0 this parameter is ignored (as all segments are assumed to be needed)

dst=

-   This is actually an option to the access std module. The path of the segment files to write. The \# characters get replaced with the segment number. So a path of "seg-###.ts" will end up with files called "seg-001.ts, seg-002.ts, seg-003.ts" etc.

index=

-   The path of the index file to write, which will contain the "playlist" of video/audio segments to stream. Recommended to end in .m3u8 by specifications. This is the file the \ tag should be pointed to.

index-url=

-   This is the URL that corresponds to the dst above (how a browser would access the dst file). The \# characters get replaced same as in the dst parameter. Note: The filename portion of this URL will most likely need to be in the exact same format as the dst parameter. So for example if dst=/www/seg-##.ts then the index-url should be something like 0http://mydomain.com/streams/seg-##.ts1 (Note the same number of \# characters)

ratecontrol=(default: false)

-   If set to false the there is no rate control (the muxer sends the data as fast/slow as it can to the streamer). If set to true the muxer should do rate-control to control the speed to muxed audio/video is sent to the streamer. Recommended setting is false in most cases. The only time I've needed this set to true is while doing a live sliding window stream of a local media file.

##### Examples

HLS Live Stream Example: Schou FishCam 0

All examples assume the following:

-   The Web Server root directory is /var/www
-   The domain name of the web server is mydomain.com
-   The stream segments & index files will be written into /var/www/streaming/ and will be accessed via 0…
-   The destination stream name index file will be called "mystream.m3u8"
-   The following HTML will allow you to view the video based on the above on an iPhone:

Re-stream a live video feed:

    % vlc -I dummy --mms-caching 0 0 vlc://quit --sout='#transcode{width=320,height=240,fps=25,vcodec=h264,vb=256,venc=x264{aud,profile=baseline,level=30,keyint=30,ref=1},acodec=mp3,ab=96}:std{access=livehttp{seglen=10,delsegs=true,numsegs=5,index=/var/www/streaming/mystream.m3u8,index-url=1,mux=ts{use-key-frames},dst=/var/www/streaming/mystream-########.ts}'

Create a VOD stream: (Non-live. When this command finishes, all the segments should have been created and the index file contain pointers to all of them)

    % vlc -I dummy /var/myvideos/video.mpg vlc://quit --sout='#transcode{width=320,height=240,fps=25,vcodec=h264,vb=256,venc=x264{aud,profile=baseline,level=30,keyint=30,ref=1},acodec=mp3,ab=96}:std{access=livehttp{seglen=10,delsegs=false,numsegs=0,index=/var/www/streaming/mystream.m3u8,index-url=0,mux=ts{use-key-frames},dst=/var/www/streaming/mystream-########.ts}'

Re-stream a live audio feed:

    % vlc -I dummy --mms-caching 0 0 vlc://quit --sout='#transcode{acodec=mp3,ab=96}:duplicate{dst=std{access=livehttp{seglen=10,delsegs=true,numsegs=5,index=/var/www/streaming/mystream.m3u8,index-url=1,mux=raw,dst=/var/www/streaming/mystream-########.mp3},select=audio}'

**Note:** I found that these example don't work as written here; I won't edit them in place as they might work in different circumstances. I found two problems using the released version of VLC 2.0.0 on WinXP:

1.  All examples need to have the single quotes surrounding the --sout parameter removed. Otherwise VLC complains "stream_out_standard stream out error: no mux specified or found by extension".
2.  The final example is audio-only, but does not specify the splitanywhere=true flag. As a result it writes one massive chunk waiting for a keyframe that never comes.

Example commandline that did work for me:

    % vlc -I dummy x:\some\audio\here.ogg vlc://quit --sout=#transcode{acodec=mp3,ab=96}:duplicate{dst=std{access=livehttp{seglen=10,splitanywhere=true,delsegs=true,numsegs=5,index=c:\temp\mystream.m3u8,index-url=0,mux=raw,dst=c:\temp\mystream-########.mp3},select=audio}

Smowton 03:23, 27 February 2012 (CET)

##### Formats supported by the iOS

See the iPhone article for a list of supported codecs as well as [Apple's HTTP Live Streaming FAQ](http://developer.apple.com/iphone/library/documentation/networkinginternet/conceptual/streamingmediaguide/FrequentlyAskedQuestions/FrequentlyAskedQuestions.html).

##### Possible improvements/fixes

from the developer:

-   Have the module auto-detect audio only streams, so the splitanywhere option is not required.
-   I'm not sure I am doing the right thing with the Win32 rename function. Linux allows me to rename a file over an existing file, even if the existing file is in use. Win32 is not so friendly. This ability is useful for updating the index file at same time it may be currently being read by the HTTP server serving the files.
-   Break the dst= and index= parameter into seperate filename/directory entries, so you only need to specify the filename format once. (instead of once for the dst= parameter, and once for the index-url= parameter)

from the community:

-   Have the module detect the codecs used and warn the end user if they are not compatible with the iPhone/iPod Touch/iPad.
-   Possible multibitrate implementation that meets the iPhone SDK specs so app developers can use VLC to host streams for their apps.

fixes:

-   Audio only streams do not validate with Apple's mediastreamvalidator. Audio streams are missing id3tags and timestamps. Check with command, "mediastreamvalidator validate --timeout=60 \[url\]"

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Streaming HowTo {#streaming-howto}

Authors:

-   Alexis de Lattre
-   Johan Bilien
-   Anil Daoud
-   Clément Stenac
-   Antoine Cellerier
-   Jean-Paul Saman


*Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.*

**This document explains how to stream, transcode and save streams using the VideoLAN solutionNOTE: See Streaming HowTo New for more updated information regarding streaming (WIP).**

#### Introduction

-   Streaming, Muxers and Codecs

#### Main

-   Easy Streaming

&nbsp;

-   Easy Streaming with Newer Versions of VLC (not completed)

&nbsp;

-   Advanced Streaming Using the Command Line

&nbsp;

-   Command Line Examples

#### VLM

-   VLM - Multiple Streaming and Video on Demand

#### Tutorials and examples

-   Receive and Save a Stream

&nbsp;

-   Convert files to other formats

&nbsp;

-   Stream a File

&nbsp;

-   Stream a DVD

&nbsp;

-   Stream a DVB Channel

&nbsp;

-   Stream from Encoding Cards and Other Capture Devices

&nbsp;

-   Stream from a DV Camcorder

&nbsp;

-   Streaming a live video feed to Darwin Streaming Server for Mobile Phones

&nbsp;

-   Streaming for the iPhone with live http

#### IPv6

-   Streaming over IPv6

&nbsp;

-   Advanced streaming with samples, multiple files streaming, using multicast in streaming

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Streaming over IPv6 {#streaming-howto-streaming-over-ipv6}

#### Streaming over IPv6

This chapter covers the specifics of streaming over IPv6. You should still read the previous chapters if you are not comfortable with streaming in general.

##### Requirements

You will obviously need an IPv6-aware operating system. That includes Windows XP/2003, Linux 2.6, Mac OS X (starting from version 10.2). Windows 2000 and Linux 2.4 are supported too, but their IPv6 stacks are not as good, so upgrade if you can. IPv6 must be properly configured and working on your system and network.

On Linux, the *ipv6* kernel module must be loaded (or compiled-in). On Windows, the IPv6 protocols suite can be installed by running "ipv6 install" from the command line, or through the Network configuration panel.

Note: Under Windows 2000, you must add by hand a default multicast IPv6 route, with the following command:

    # ipv6 rtu ff::/8 4

where the last number (*4* in this example) is the number of your true IPv6 interface. To have a list of your IPv6 interfaces, run **ipv6 if**

Warning: Under Windows XP SP1, you may have problems with a hidden IPv6 firewall. To solve the problem, go to the list of Windows Services and stop the IPv6 firewalling service. You should consider upgrading to Service Pack 2 which provides an integrated IPv4/IPv6 firewall that can be configured through the GUI.

Warning: If you are using VMware under Linux, you will have to stop VMware and unload the VMware kernel modules, because we noticed it prevented IPv6 streaming!

##### Limitations

There are still some features of the VLC media player which do not support IPv6. In particular, it is not possible to use RTSP over IPv6 because the underlying library, Live.com, does not support IPv6 at the time of writing.

Additionally, note that at the moment, VLC defaults to using IPv4 mostly every, as it is what most people uses. That might be changed to something more transparent in future versions.

##### Streaming with VLC

###### With the Streaming Wizard (GUI)

The streaming wizard accepts IPv6 addresses between braces, for example: **\[2002:8ac3:802d:1242:211:11ff:fe25:e6b4\]**. If you specify a link-local address, you will most likely need to specify the networking interface to use. On Unix, that can be done this way: **\[fe80::211:11ff:fe25:e6b4%eth0\]** to attach to eth0. Similarly, on Windows, you may specify **\[fe80::211:11ff:fe25:e6b4%1\]** where 1 is the number of the network interface as defined by the operating system.

If you're streaming over HTTP, note that IPv6 is automatically used by default (so that both IPv6 and IPv4 clients will be allowed).

If you want to specify DNS hostname, keep in mind that the VLC defaults to IPv4 resolution. You must either specify hostnames that only resolves to IPv6 addresses, or enable the "Force IPv6" *advanced* option in *Preferences / General Settings / Input*.

###### From the command-line

The **--ipv6** command line option force the use of IPv6 by default (ie. IPv6 is always attempted before IPv4).

    % vlc -vvv video1.xyz --ipv6 --sout udp:[ff08::1] --ttl 12

where:

-   **video1.xyz** is the file you want to stream (you can also put **dvdsimple:/dev/dvd** to stream a DVD or any other input configuration),
-   **ff08::1** is either:
    -   the IPv6 address of the machine you want to unicast to;
    -   or the multicast IPv6 address.
-   **12** is the value of the TTL (Time To Live) of your IP packets (which means that the stream will be able to cross 11 routers).

Note: Under Unix/Linux, you may have to protect the square brackets around the IPv6 address:

    % vlc -vvv video1.xyz --ipv6 --sout udp:\[ff08::1\] --ttl 12

Note: You may have to specify the output network interface:

    % vlc -vvv video1.xyz --ipv6 --sout udp:[ff08::1%eth0] --ttl 12

where **eth0** is the name of the network interface (under Linux the network interfaces are named **ethX**, under Mac OS X it's **enX** and under Windows it's **X**, where **X** is the appropriate number).

##### Receiving an IPv6 stream

###### With the graphical user interface

Select File / Open Network Stream. To receive an UDP/RTP unicast stream sent to your system, you should select the Force IPv6 option (and possibly adjust the destination UDP port). To receive an UDP multicast stream, select the UDP/RTP Multicast option, and specify the multicast address to subscribe to inside square brackets. The IPv6 addresses syntax is the same as that explained in the *Streaming over IPv6* section of this chapter.

###### From the command line

As for streaming, the **--ipv6** command line option force the use of IPv6 by default (i.e., IPv6 is always attempted before IPv4).

    % vlc -vvv --ipv6 udp:@[ff08::1]

Under Unix/Linux, you may have to protect the square brackets around the IPv6 address:

    % vlc -vvv --ipv6 udp:@\[ff08::1\]

You may have to specify the output network interface:

    % vlc -vvv video1.xyz --ipv6 --sout udp:[ff08::1%eth0] --ttl 12

where **eth0** is the name of the network interface (under Linux the network interfaces are named **ethX**, under Mac OS X it's **enX** and under Windows it's **X**, where **X** is the appropriate number).

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### VLM {#streaming-howto-vlm}

This page is outdated and information might be incorrect.

#### VLM

*VideoLAN Manager* is a small media manager designed to control multiple streams with *only one instance of VLC*. It allows multiple streaming and video on demand (VoD). This manager being a new feature, it can only be controlled by the telnet interface or the http interface.

##### Interfaces

###### Telnet interface

You can launch the telnet interface as a common interface using the command line:

    % vlc --intf telnet

    % vlc --extraintf telnet

The telnet interface can also be launched in the wxWindows interface:

The default port is 4212. The default password is "admin". These can be changed using **--telnet-port \** and **--telnet-password \** command line options. They can also be changed in the preferences panel when using the wxWindows interface in the *Modules*-\>*interface*-\>*telnet* section (check the *Advanced options* checkbox).

###### HTTP interface

Launching the HTTP interface is described in Control_VLC_via_a_browser

To access the vlm section of the http interface, use the following URL: *0* (*1* for VLC 0.8.4 and older).

**Note:** People who aren't used to command line streaming with VLC but want to use VLM's features are advised to use the HTTP interface.

##### VLM Elements

###### Medias

A *Media* is composed with a list of inputs (the video and audio streams you want to stream), an output (how and where you want to stream them) and some options.

There are two types of medias:

-   **vod**: A vod media is commonly used for Video on Demand. It will be launched only if a vod client asks for it.
-   **broadcast**: A broadcast media is very close to a TV program or channel. It is launched, stopped or paused by the administrator and may be repeated several times. The client has no control over this media.

###### Schedules

A *Schedule* is a script with a date. When the schedule date is reached, the script is launched. There are several options available like a period or a number of repetitions.

##### Command line syntax

###### Command lines

-   **help**: Displays an exhaustive command lines list
-   **new (name) vod\|broadcast\|schedule \[properties\]**: Create a new vod, broadcast or schedule element. Element names must be unique and cannot be "media" or "schedule". You can specify properties in this command line or later on by using the **setup** command.
-   **setup (name) (properties)**: Set an elements property. See #Media Properties.
-   **show \[(name)\|media\|schedule\]**: Display current element states and configurations.
    -   **show (name)**: Specify an element's name to show all information concerning this element.
    -   **show media** displays a summary of media states.
    -   **show schedule** displays a summary of schedule states.
-   **del (name)\|all\|media\|schedule**: Delete an element or a group of elements. If the element wasn't stopped, it is first stopped before being deleted.
    -   **del (name)**: Delete the (name) element.
    -   **del all**: Delete all elements
    -   **del media**: Delete all media elements.
    -   **del schedule**: Delete all schedule elements.
-   **control (name) \[instance_name\] (command)**: Change the state of the (instance_name) instance of the (name) media. If (instance_name) isn't specified, the control command affects the default instance. See #Control Commands for available control commands.
-   **save (config_file)**: Save all media and schedule configurations in the specified config file. The config file path is relative to the directory in which vlc was launched. If the file exists it will be overwritten. Note that states, such as playing, paused or stop, are not saved. See #Configuration Files for more info.
-   **load (config_file)**: Load a configuration file. The config file path is relative to the directory in which vlc was launched. See #Configuration Files for more info.

###### Media Properties

Note: Except the "append" property, all properties can be followed by another one.

-   **input (input_name)**: Add an input to the end of the media's input list.
-   **output (output_name)**: Define the media's output. The syntax is the same as the vlc ":sout=..." vlc option but you do not have to put the ":sout=..." string. See Documentation:Streaming HowTo/Advanced Streaming Using the Command Line for more information concerning stream outputs (sout). Note: You do not have to specify an output for vod elements.
-   **option (option_name)\[=value\]** : Adds the (option_name) to the media option list. The syntax is equivalent to the ":(option)=..." option , but you do not have to put the ":" string. Options are global: they are applied to all inputs of the media.
-   **enabled\|disabled**: Enable or Disable the media. If a media is disabled, it cannot be streamed, paused, launched by a schedule, or played as VoD.
-   **loop\|unloop (broadcast only)**: If a media with the "loop" option receives the "play" command, it will automatically restart to play the input list once the end of the input list is reached. Note: **loop\|unloop** is only used for broadcast media types.
-   **mux (mux_name)**: This option should only be specified if you want the elementary streams to be sent encapsulated instead of raw. The (mux_name) should be specified as a four characters length identifier such as mp2t for MPEG TS or mp2p for MPEG PS. See Documentation:Streaming HowTo/Streaming, Muxers and Codecs. Note: The **mux** property is only used for vod media types.

###### Schedule Properties

-   **enabled\|disabled**: A disabled schedule will never be launched.
-   **append (command_until_rest_of_the_line)**: Add a command to the command line lit. The command line can be every command VLM can understand. Note: The rest of the line will be considered as part of the command line. You cannot put another option after the **append** one.
-   **date (year)/(month)/(day)-(hour):(minutes):(seconds)\|now**: Specify the first date the schedule should be launched. You can specify a date using the **(year)/(month)/(day)-(hour):(minutes):(seconds)** format (example: 2004/11/16-00:43:12) or using the **now** keyword. If **now** is used, the schedule will be launched as soon as possible (i.e. as soon as it is enabled) and the current date will be used as the first date of the schedule.
-   **period (years_aka_12_months)/(months_aka_30_days)/(days)-(hours):(minutes):(seconds)**: Specify the period of time a schedule must wait for launching itself another time. (Months are considered as 30 days, Years as 12 months) If a period is specified without a **repeat** property, the schedule will be launched endlessly.
-   **repeat (number_of_repetitions)**: Specify the number of times the schedule will be launched again. For example, if a schedule has **repeat 11** it will be launched 12 times.

###### Control Commands

-   **play**: Stat a broadcast media. The media begins to launch the first item of the input list, then launches the next one and so on. (like a play list)
-   **pause**: Put the broadcast media in paused status.
-   **stop**: Stop the broadcast media.
-   **seek (percentage)**: Seek in the current playing item of the input list.

##### Configuration Files

A VLM configuration file is a list of command lines : one line corresponds to one command line.

To create a configuration file, just edit a text file and type a list of VLM commands. Beware of recursive calls: you can put a **load (file)** in a configuration file which can lead to recursive inclusion of the same file and result in VLC's crash.

You can automatically load a VLM configuration when launching VLC with the --vlm-conf \ command line option. The minimal command to make that work is:

    % vlc -I telnet --vlm-conf vlm.conf

As of versions \> 0.8.1, any line where the first non whitespace character is a \# is considered as a comment.

#### Examples

This section provides several small vlm configuration files.

##### Multiple streams

###### Simple broadcasting

    new channel1 broadcast enabled
    setup channel1 input 0
    setup channel1 output #rtp{mux=ts,dst=239.255.1.1,port=5004,sdp=sap://,name="Channel 1"}

    new channel2 broadcast enabled
    setup channel2 input udp://@239.255.12.42
    setup channel2 output #rtp{mux=ts,dst=239.255.1.2,port=5004,sdp=sap://,name="Channel 2"}

    control channel1 play
    control channel2 play

-   if you are using direct show and are getting "control : unknown error" try "setup *channel* enabled"

###### Scheduled broadcasting

    new my_media broadcast enabled
    setup my_media input my_video.mpeg input my_other_movie.mpeg
    setup my_media output #rtp{mux=ts,dst=239.255.1.1,sdp=sap://,name="My Media"}

    new my_sched schedule enabled
    setup my_sched date 2012/12/12-12:12:12
    setup my_sched append control my_media play

##### Video On Demand

###### Basic example

First launch the vlc

    % vlc --ttl 12 -vvv --color -I telnet --telnet-password videolan --rtsp-host 0.0.0.0 --rtsp-port 554

where:

-   **12** is the value of the TTL (Time To Live) of your IP packets (which means that the stream will be able to cross 11 routers).
-   **telnet** launches the telnet interface of the vlc.
-   **videolan** is the password to connect to the telnet interface.
-   **0.0.0.0** is the host address.
-   **554** is the port on which you stream.

Then you connect to the vlc telnet interface and create the vod object. You can connect to vlc telnet interface by use the terminal.

    % telnet localhost 4212

and create the vod object.

    new Test vod enabled
    setup Test input my_video.mpg

You can access to the stream with:

    % vlc rtsp://server:554/Test

where:

-   **server** is the address of the streaming server (IP or DNS)

###### Advanced example

You can also specify options, a muxer, or an additional output chain that will be prepended to the RTP output used by VoD (e.g. to enable transcoding).

**Note:** make sure to enter the corresponding commands before the VoD media is enabled, or before you setup the input.

    new Test2 vod
    setup Test2 output #transcoding{vcodec=h264,vb=512,acodec=mp4a,ab=96}
    setup Test2 mux mp2t
    setup Test2 input my_video.mpg
    setup Test2 enabled

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

## Troubleshooting {#troubleshooting}

### Basic troubleshooting {#basic-troubleshooting}

**Languages:English** • français

#### VLC Support Guide: Solve your VLC issues right now!

The **V**LC **S**upport **G**uide is an informal, step-by-step guide for troubleshooting most common issues with VLC.

It complements the VLC media player Documentation.

**So what's your problem?**

##### Installation Issue

VLC won't install. Go!

##### Startup Issue

VLC won't start up. Go!

##### Audio Playback Issue

The audio or the sounds are wrong. Go!

##### Video Playback Issue

The video is messed up. Go!

##### Subtitle Display Issue

The subtitles aren't working properly. Go!

##### Usage Issue

I have difficulty using VLC. Go!

##### Interface Issue

I want to change my interface. Go!

##### Uninstallation Issue

VLC won't uninstall (why are you uninstalling it anyway?). Go!

#### Get Help

If this troubleshooter does not resolve your problems or answer your questions, some other resources which you can use include:

-   Frequently asked questions
-   Frequently asked questions about VLC on Windows
-   Frequently asked questions about VLC on macOS
-   Frequently asked questions about VLC on Linux
-   The [VideoLAN support forum](https://forum.videolan.org/)
-   The VideoLAN IRC channel.
-   VLC documentation

This page is part of the informal VLC Support Guide.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Misc {#misc}

Miscellaneous other cool things you can do with VLC.

#### Snapshot Tool

Did you know you can use special codes to automatically generate filenames in the Snapshot Tool?

#### Audio Bar Graph over Video

This section specifies how to enable the audiobargraph audio filter and video overlay, (mostly) via the GUI. This displays an audio meter overlaid on the video.

There are three parts: **an audio filter**, which sends its output via TCP to the **Remote Control** (RC) Interface. This information is then picked up and displayed by the **Audio Bar Graph video subpicture filter** (OSD).

To enable this, VLC needs to be started with the **0** command-line switch e.g.

    > "%PROGRAMFILES%\VideoLAN\VLC\vlc.exe" --rc-host localhost:12345

In the GUI, set the following (this example from VLC v1.1.9 on Windows 7):

-   Preferences:Show settings:All
-   **Audio → Filters** Enable "Audio part of the BarGraph function"
-   **Audio → Filters → Audiobar Graph** Use defaults, change "Sends the barGraph information every n audio packets" to 1 to enable viewing a more accurate display
-   **Interface → Main interfaces** Enable "Remote control interface"
-   --**Interface → Main interfaces → RC** Enable "Do not open a DOS command box interface"^\[Check\ this\ —\ outdated ?\]^
-   **Video → Subtitles-OSD** Enable "Audio Bar Graph Video sub filter"
-   **Video → Subtitles-OSD → Audio Bar Graph** Set the following settings:
    -   **Value of the audio channels levels = 0** (setting this to 0:1 crashes VLC v1.1.9)
    -   **X coordinate = 0**
    -   **Y coordinate = 0** (this doesn't seem to affect anything)
    -   **Transparency of the bargraph = 128** for 50% transparency which looks ok
    -   **Bargraph position = Left** (seems to only work Left,Center,Right—can't go top or bottom)
    -   **Alarm = 1** (enables the silence alarm: puts a red border around the bargraph if silent for too long)
    -   **Bar width in pixel = 10** (20 if you want it to be really visible)

#### How to show album art

1.  Close the VLC media player.

2.  Open the config file (Windows: 0) with your preferred text editor.

3.  Search for:

        #metadata-network-access=0

4.  Delete the leading hash sign and change **0** to **1**.

5.  Save the file.

6.  Open a MP3- or a M3U-File in a folder there a file like Folder.jpg, AlbumArtSmall.jpg, AlbumArt.jpg, Album.jpg, .folder.png, cover.jpg or thumb.jpg exists. The player should show the album art.

If you have none of the files above but any other file you can put this filename in vlcrc. Search for **0**.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

## Using VLC {#using-vlc}

### Alternative Interfaces {#alternative-interfaces}

This page is outdated and information might be incorrect.

#### The HTTP interface

VLC ships with a little HTTP server integrated. It is used both to stream using HTTP, and for the HTTP remote control interface.

To start VLC with the HTTP interface, use:

    % vlc -I http [--http-src /directory/] [--http-host host:port]

If you want to have both the "normal" interface and the HTTP interface, use **vlc --extraintf http**.

The HTTP interface will start listening at host:port (\:8080 if omitted), and will reproduce the structure of /directory at http://host:port/ ( vlc_source_path/share/http if omitted ).

Use a browser to go to http://your_host_machine:port. You should be taken to the main page.

VLC is shipped with a set of files that should be enough for generic needs. It is also possible to customize pages. See Documentation:Play HowTo/Building Pages for the HTTP Interface.

Available pages for 1.0.3 :

-   http://host:port - Main Interface
-   0 - VLM Interface
-   0 - Mosaic Wizard
-   0 - Flash based remote playback

#### Ncurses

This is a text interface, using ncurses library.

Start VLC with **-I ncurses** or **--extraintf ncurses**. You will then get something like that:

The ncurses interface

Press h to get the list of all available commands, with a short description.

There is also a filebrowser available for the ncurses interface in order to add playlist items. Press 'B' to use it.

The ncurses filebrowser

You can set the filebrowser starting point by launching vlc with the **--browse-dir** option:

    % vlc -I ncurses --browse-dir /filebrowser/starting/point/

#### Other control interfaces

VLC includes a number of so-called interfaces that are not really interfaces, but means of controlling VLC. Nevertheless, they are enabled by setting them as interface or extra interface, either in the Preferences, in General/Interface, or using **-I** or **--extraintf** on the command line.

##### Hotkeys

This module allows you to control VLC and playback via hotkeys. It is always enabled by default. You can use hotkeys in the video output window, you can't in the audio dummy interface.

Hotkeys can be hacked by:

    % vlc --key-

Code is composed by modifiers keys (Alt, Shift, Ctrl, Meta,Command) separated by a dash (-) and terminated by a key (a...z, +, =, -, ',', +, \<, \>, \`, /, ;, ', \\, \[, \], \*, Left, Right, Up, Down, Space, Enter, F1...F12, Home, End, Menu, Esc, Page Up, Page Down, Tab, Backspace, Mouse Wheel Up and Mouse Wheel Down). Main controls are available from hotkeys, such as : fullscreen, play-pause, faster, slower, next, prev, stop, quit, vol-up, etc. (use the **--longhelp** option for full list of functions). For example, for binding fullscreen to Ctrl-f, run:

    % vlc --key-fullscreen 'Ctrl-f'

The list of the default hotkeys is available here.

##### RC, Telnet

These two interfaces allow you to control VLC from a command shell (possibly using a remote connexion or a Unix socket).

Start VLC with **-I rc** or **--extraintf rc**. When you get the **Remote control interface initialized, \`h' for help** message, press h and Enter to get help about available commands.

To be able to remote connect to your VLC using a TCP socket (telnet-like connexion), use **--rc-host your_host:port**. Then, by connecting (using telnet or netcat) to the host on the given port, you will get the command shell.

To use a UNIX socket (local socket, this does not work for Windows), use **--rc-unix /path/to/socket**. Commands can then be passed using this UNIX socket.

The RTCI interface is an old module merged into the RC interface.

##### Gestures

Gestures provide a simple mouse gestures control. TODO

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Command line {#command-line}

See also: Command-line interface

This page is outdated and information might be incorrect.

#### Use the command line

**TODO: completely outdated**

All standard operations of VLC should be available from the GUI. However, some complex operations can only be done from the command line and there are situations in which you don't need or want a GUI. Here is the complete description of VLC's command line and how to use it.

You need to be quite comfortable with command line usage to use this.

    Note: Windows users have to use the --option-name="value" syntax instead of the --option-name value syntax.

#### Getting help

VLC uses a modular structure. The core mainly manages communication between [modules](#modules). All the multimedia processing is done by modules. There are input modules, demultiplexers, decoders, video output modules, ...

This chapter will only describe the "general" options, i.e., the core options. Each module adds new options. For example, the HTTP input module will add options for caching, proxy, authentication, ...

By using **vlc --help**, you will get the basic core options. **vlc --longhelp** will give all the basic options (core + modules). Adding **--advanced** will give the "advanced options" (for advanced users). So **vlc --longhelp --advanced** will give you all options. You can also append **--help-verbose** if you want more detailed help.

Also, you might want to get debug informations. To do this, use **-v** or **-vv** (this will show lower severity messages). If your console supports it, you can add **--color to get messages in color.**

#### Opening streams

The following commands start VLC and start reading the given element(s):

##### Opening a file

Start VLC with:

    % vlc my_file

VLC should be able to recognize the file type. If it does not, you can force demultiplexer and decoder (see below).

A list of all video and audio codecs supported by VLC is available on the [VLC features list](http://www.videolan.org/vlc/features.html).

##### Opening a DVD or VCD, or an audio CD

Start VLC with:

For a DVD with menus:

    % vlc dvd://[device][@raw_device][@[title][:[chapter][:angle]]]

In most cases, **vlc dvd://** or **vlc dvd://\[device\]** will do.

-   On GNU/Linux \[device\] is the path to the block device: e.g., **vlc dvd:///dev/dvd**.
-   On Windows, \[device\] is the drive letter with **/** and **:/**: e.g., **vlc dvd:///D:/**.

or

(DVD without menus):

    % vlc dvdsimple://[device][@raw_device][@[title][:[chapter][:angle]]]

or

(VCD):

    % vlc vcd://[device][@{E|P|E|T|S}[number]]

or

(Audio CD):

    % vlc cdda://[device][@[track]]

##### Receiving a network stream

To receive an unicast RTP/UDP stream (sent by VLC's stream output), start VLC with:

    % vlc rtp://@:5004

If 5004 is the port to which packets are sent. 1234 is another commonly used port number. you use the default port (1234), **vlc rtp://** will do. For more information, look at the Streaming Howto.

To receive an multicast UDP/RTP stream (sent by VLC's stream output), start VLC with:

    % vlc rtp://@multicast_address:port

To receive a SSM (source specific multicast) stream, you can use:

    % vlc rtp://server_address@multicast_address:port

This only works on OSs that support SSM (Windows XP and Linux).

To receive a HTTP stream, start VLC with:

    % vlc 0

To receive a RTSP stream, start VLC with:

    % vlc rtsp://www.example.org/your_stream

#### Modules selection

See also: [Documentation:Modules](#modules)

VLC always tries to select the most appropriate interface, input and output modules, among the ones available on the system, according to the stream it is given to read. However, you may wish to force the use of a specific module with the following options.

-   **--intf \** allows you to select the interface module.
-   **--extraintf \** allows you to select extra interface modules that will be launched in addition to the main one. This is mainly useful for special *control* interfaces, like HTTP, RC (Remote Control), ... (see below)
-   **--aout \** allows you to select the audio output module.
-   **--vout \** allows you to select the video output module.
-   **--memcpy \** allows you to choose a memory copy module. You should probably never touch that.

You can get a listing of the available modules by using **vlc -l**

#### Stream Output

The Stream output system allows vlc to become a streaming server.

For more details on the stream output system, please have a look at the [Streaming HowTo](#streaming-howto).

#### Other Options

##### Audio options

Note that in recent versions (3.x.x branch, possibly earlier):

-   0 no longer exists: use 1 instead
-   0 no longer exists but 1 and 2 may be used
-   0 no longer exists: 1 might be equivalent?
-   0 no longer exists: use 1 instead

------------------------------------------------------------------------

-   **--audio**, **--no-audio** disables audio output. Note that if you are streaming (ex: to a file) this has no effect (streaming copies the audio verbatim). Use --sout-xxx instead (ex: --no-sout-audio)
-   **--gain \** audio gain (between 0 and 8)
-   **--volume-step \** audio output volume step (between 1 and 256)
-   **--volume-save**, **--no-volume-save** remember the volume (default enabled)
-   **--spdif**, **--no-spdif** Force S/PDIF support (default disabled)
-   **--force-dolby-surround** {0 (Auto), 1 (On), 2 (Off)} Force detection of Dolby Surround
-   **--stereo-mode** {0 (Unset), 1 (Stereo), 2 (Reverse stereo), 3 (Left), 4 (Right), 5 (Dolby Surround), 6 (Headphones)} Stereo audio output mode
-   **--audio-desync \** Audio desynchronization compensation
-   **--audio-replay-gain-mode** {none,track,album} Replay gain mode
-   **--audio-replay-gain-preamp \** Replay preamp
-   **--audio-replay-gain-default \** Default replay gain
-   **--audio-replay-gain-peak-protection**, **--no-audio-replay-gain-peak-protection** Peak protection (default enabled)
-   **--audio-time-stretch**, **--no-audio-time-stretch** Enable time stretching audio (default enabled)
-   **-A**, **--aout** {any,pulse,alsa,sndio,adummy,afile,amem,none} Audio output module
-   **--role** {video,music,communication,game,notification,animation,production,accessibility,test} Media role
-   **--audio-filter \** adds audio filters to the processing chain. Available filters are visual (visualizer with spectrum analyzer and oscilloscope), headphone (virtual headphone spatialization) and normalizer (volume normalizer)
-   **--audio-visual** {any,visual,glspectrum,none} Audio visualizations
-   **--audio-resampler** {any,samplerate,ugly,soxr,speex_resampler,none} Audio resampler

##### Video options

-   **--no-video** disables video output.
-   **--grayscale** turns video output into grayscale mode.
-   **--fullscreen** ( or **-f**) sets fullscreen video.
-   **--nooverlay** disables hardware acceleration for the video output.
-   **--width, --height \** sets the video window dimensions. By default, the video window size will be adjusted to match the video dimensions.
-   **--start-time \** starts the video here; the integer is the number of seconds from the beginning (e.g. 1:30 is written as 90)
-   **--stop-time \** stops the video here; the integer is the number of seconds from the beginning (e.g. 1:30 is written as 90)
-   **--zoom \** adds a zoom factor.
-   **--aspect-ratio \** forces source aspect ratio. Modes are 4x3, 16x9, ...
-   **--spumargin \** forces SPU subtitles postion.
-   **--video-filter \** adds video filters to the processing chain. You can add several filters, separated by commas
-   **--video-splitter \** adds video splitters to the processing chain. (wall, panoramix, clone)
-   **--sub-filter \** adds video subpictures filter to the processing chain.

##### Desktop/Screen grab options

You can see the various options for "grabbing the desktop" (VLC's built-in screen grabber capture device) by using the GUI. See 0

##### Playlist options

-   **--random** plays files randomly forever.
-   **--loop** loops playlist on end.
-   **--repeat** repeats current item until another item is forced
-   **--play-and-stop** stops the playlist after each played item.
-   **--no-repeat --no-loop** prevents the video from being executed again. (Useful when want to encode a file)

##### Network options

-   **--server-port \** sets server port.
-   **--iface \** specifies the network interface to use.
-   **--iface-addr \** specifies your network interface IP address.
-   **--mtu \** specifies the MTU of the network interface.
-   **--ipv6** forces IPv6.
-   **--ipv4** forces IPv4.

##### CPU options

You should probably not touch these options unless you know what you are doing.

-   **--nommx** disables the use of MMX CPU extensions.
-   **--no3dn** disables the use of 3D Now! CPU extensions.
-   **--nommxext** disables the use of MMX Ext CPU extensions.
-   **--nosse** disables the use of SSE CPU extensions.
-   **--noaltivec** disables the use of Altivec CPU extensions.

##### Miscellaneous options

-   **--quiet** deactivates all console messages.
-   **--color** displays color messages.
-   **--search-path \** specifies interface default search path.
-   **--plugin-path \** specifies plugin search path.
-   **--no-plugins-cache** disables the plugin cache (plugins cache speeds up startup)
-   **--dvd \** specifies the default DVD device.
-   **--vcd \** specifies the default VCD device.
-   **--program \** specifies program (SID) (for streams with several programs, like satellite ones).
-   **--audio-type \** specifies the default audio type to use with dvds.
-   **--audio-channel \** specifies the default audio channel to use with dvds.
-   **--spu-channel \** specifies the default subtitle channel to use with dvds.
-   **--version** gives you information about the current VLC version.
-   **--module \** displays help about specified module. (Shortcut: **-p**)

#### Item-specific options

There are many options that are related to items (like **--novideo**, **--codec**, **--fullscreen**).

For all of these, you have the possibility to make them item-specific, using ":" instead of "--" and putting the option just after the concerned item.

Examples:

    % vlc file1.mpg :fullscreen file2.mpg

will play file1.mpg in fullscreen mode and file2.mpg in the default mode (which is generally no fullscreen), whereas

    % vlc --fullscreen file1.mpg file2.mpg

will play both files in fullscreen mode

    % vlc --fullscreen file1.mpg :sub-file=file1.srt :no-fullscreen file2.mpg :filter=distort

will play file1.mpg in windowed (no-fullscreen) mode with the subtitles file file1.srt and will play file2.mpg with video filter distort enabled in fullscreen mode (item-specific options override global options).

#### Filters

These are the old style VLC filters. They only apply to on screen display and thus cannot be streamed. However, on version 1.1.11 you are still able to apply these filters in *transcode* module using parameter *vfilter*. More information can be found on Advanced Streaming Using the Command Line.

##### Deinterlacing video filter

*Further information: Documentation:Modules/deinterlace*

Module name: **deinterlace**

-   **sout-deinterlace-mode \ {discard,blend,mean,bob,linear,x,yadif,yadif2x,phosphor,ivtc}** : Streaming deinterlace mode. Deinterlace method to use for streaming
-   **sout-deinterlace-phosphor-chroma \ {1,2,3,4}** : Phosphor chroma mode for 4:2:0 input. Choose handling for colours in those output frames that fall across input frame boundaries.
    -   Latest (1): take chroma from new (bright) field. Good for interlaced input, such as videos from a camcorder
    -   AltLine (2): take chroma line 1 from top field, line 2 from bottom field, etc. Default, good for NTSC telecined input (anime DVDs, etc.)
    -   Blend (3): average input field chromas. May distort the colours of the new (bright) field, too
    -   Upconvert (4): output in 4:2:2 format (independent chroma for each field). Best simulation, but requires more CPU and memory bandwidth *default value: 2*
-   **sout-deinterlace-phosphor-dimmer \ {1,2,3,4}** : Phosphor old field dimmer strength: 1 (Off), 2 (Low), 3 (Medium), 4 (High). This controls the strength of the darkening filter that simulates CRT TV phosphor light decay for the old field in the Phosphor framerate *default value: 2*

##### Invert video filter

*Further information: Documentation:Modules/invert*

Module name: **invert**

##### Image properties filter

*Further information: Documentation:Modules/adjust*

Module name: **adjust**

-   **contrast \** : Contrast *default value: 1.0*
-   **brightness \** : Brightness *default value: 1.0*
-   **hue \** : Hue *default value: 0*
-   **saturation \** : Saturation *default value: 1.0*
-   **gamma \** : Gamma *default value: 1.0*
-   **brightness-threshold \** : When this mode is enabled, pixels will be shown as black or white. Also may invert the brightness value. The threshold value will be the brightness defined below *default value: disabled*

##### Wall video filter

*Further information: Documentation:Modules/wall*

Module name: **wall**

This filter splits the output in several windows.

-   **wall-cols \** : Number of horizontal windows in which to split the video *default value: 3*
-   **wall-rows \** : Number of vertical windows in which to split the video *default value: 3*
-   **wall-active \** : Comma-separated list of active windows, defaults to all *default value: NULL*
-   **wall-element-aspect \** : Aspect ratio of the individual displays building the wall *default value: 4:3*

Note: for 0, to select windows 2 and 4 you would write **--wall-active 2,4**. When this option isn't specified, all windows are displayed.

##### Video transformation filter

*Further information: Documentation:Modules/transform*

Module name: **transform**

-   **transform-type \ { "90", "180", "270", "hflip", "vflip", "transpose", "antitranspose" }** : Transformation type *default value: "90"*

##### Distort video filter

*Further information: Documentation:Modules/distort*

Module name: **distort**

##### Clone video filter

*Further information: Documentation:Modules/clone*

This filter clones the output window.

Module name: **clone**

-   **clone-count \** : Number of video windows in which to clone the video. *default value: 2*
-   **clone-vout-list \** : You can use specific video output modules for the clones. Use a comma-separated list of modules. *default value: ""*

##### Crop video filter

*Further information: Documentation:Modules/crop*

Module name: **crop**

-   **crop-geometry \** : Set the geometry of the zone to crop. This is set as 0
-   **autocrop \** : Automatically detect black borders and crop them *default value: disabled*
-   **autocrop-ratio-max \** : Maximum image ratio. The crop plugin will never automatically crop to a higher ratio (ie, to a more "flat" image). The value is ×1000: 1333 means 4⁄3 *default value: 2405*
-   **crop-ratio \** : Force a ratio (0 for automatic). Value is ×1000: 1333 means 4⁄3 *default value: 0*
-   **autocrop-time \** : The number of consecutive images with the same detected ratio (different from the previously detected ratio) to consider that ratio changed and trigger recrop *default value: 25*
-   **autocrop-diff \** : The minimum difference in the number of detected black lines to consider that ratio changed and trigger recrop *default value: 16*
-   **autocrop-non-black-pixels \** : The maximum of non-black pixels in a line to consider that the line is black *default value: 3*
-   **autocrop-skip-percent \** : Percentage of the line to consider while checking for black lines. This allows skipping logos in black borders and crop them anyway *default value: 17*
-   **autocrop-luminance-threshold \** : Maximum luminance to consider a pixel as black (0-128) *default value: 40*

##### Motion blur filter

*Further information: Documentation:Modules/motionblur*

Module name: **motionblur**

-   **motionblur-factor \** : The bluring factor (1 to 127). Higher values mean more blurring *default value: 80*

##### Video pictures blending

*Further information: Documentation:Modules/blend*

Module name: **blend**

##### Video scaling filter

*Further information: Documentation:Modules/scale*

Module name: **scale**

#### Subpictures Filters

These are the new VLC filters. They can be streamed.

##### Marquee display sub filter

*Further information: Documentation:Modules/marq*

Module name: **marq**

-   **marq-marquee \** : Marquee text to display. *default value: VLC*
-   **marq-file \** : File to read the marquee text from. *default value: NULL*
-   **marq-x \** : X offset, from the left screen edge. *default value: 0*
-   **marq-y \** : Y offset, down from the top. *default value: 0*
-   **marq-position \** : Marquee position: 0=center, 1=left, 2=right, 4=top, 8=bottom, you can also use combinations of these values, eg 6 = top-right. *default value: -1*
-   **marq-opacity \** : Opacity (inverse of transparency) of overlaid text. 0 = transparent, 255 = totally opaque. *default value: 255*
-   **marq-color \ { 0x000000, 0x808080, 0xC0C0C0, 0xFFFFFF, 0x800000, 0xFF0000, 0xFF00FF, 0xFFFF00, 0x808000, 0x008000, 0x008080, 0x00FF00, 0x800080, 0x000080, 0x0000FF, 0x00FFFF }** : Color of the text that will be rendered on the video. This must be an hexadecimal (like HTML colors). The first two chars are for red, then green, then blue. *default value: 0xFFFFFF*
-   **marq-size \** : Font size, in pixels. 0 uses the default font size. *default value: 0*
-   **marq-timeout \** : Number of milliseconds the marquee must remain displayed. 0 means forever. *default value: 0*
-   **marq-refresh \** : Number of milliseconds between string updates. This is mainly useful when using meta data or time format string sequences. *default value: 1000*

The time sub filter was merged into this module.

##### Logo video filter

*Further information: Documentation:Modules/logo*

Module name: **logo**

This filter can be used both as an old style filter or a subpictures filter.

-   **logo-file \** : Image to display. The full format is 0.
-   **logo-x \** : X offset from upper left corner. *default value: 0*
-   **logo-y \** : Y offset from upper left corner. *default value: 0*
-   **logo-position \ { 0, 1, 2, 4, 8, 5, 6, 9, 10 }** : Logo position. *default value: 5*
-   **logo-opacity \** : Logo opacity. 0 is transparent, 255 is fully opaque. *default value: 255*
-   **logo-delay \** : Global delay in [ms](http://en.wiktionary.org/wiki/ms#Translingual). Sets the duration each image will be displayed for in a loop iteration unless specified otherwise in the 0 option. *default value: 1000*
-   **logo-repeat \** : Number of loops for the logo animation. -1 for continuous, 0 to disable. *default value: -1*

Note: You can move the logo by left-clicking on it.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Hotkeys {#hotkeys}

Most of VLC functions are accessible using hotkeys.

The list of the available hotkeys and their functions can be retrieved and altered in the *Preferences* panel of the player. In the Windows and Linux interface, *Preferences* are available in the "Tools" tab as the "Preferences" menu item. In the MacOS X interface, open the "VLC" menu, and select "Preferences". Select the "Hot keys" panel in the dialog.

As of version 0.9, a list of hotkeys are presented in a drop-down window. To change one, double-click its name to select it. Then, press the new key that will trigger the specified action. Modifier keys (such as Control/Command and Alt) may also be used. In the 1.x version you can also filter hotkeys with a search filter.

In earlier versions, several boxes gave the list of modifiers for the hotkey. To trigger an action using a hotkey, you need to press simultaneously the keys corresponding to the different selected modifiers as well as the key set in the dropdown.

To change the binding of a hotkey, select or deselect boxes corresponding to the different modifiers, and change the key by using the drop-down menu. Select the *Save* button to apply the changes.

The Hotkeys Panel - MacOS X interface**FIXME - needs verifying for 0.9**

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Interface {#interface}

#### General Interface Description

VLC has several interfaces:

-   A cross-platform interface for Windows and GNU/Linux, which is called Qt.
-   A native Mac OS X interface.
-   An interface that supports skins for both Windows and GNU/Linux.

The operation of VLC is essentially the same in all the interfaces.

##### Windows and GNU/Linux (Qt)

The screenshot below shows the default interface in VLC 2.0. More features can be displayed by selecting them in the *View* menu.

See also VLC Interface 2.0 on Windows 7

##### Mac OS X

This screenshot shows the default interface that VLC had on Mac OS X until version 1.1:

Since version 2.0 the interface has been redesigned. See OSX 2.0 interface.

#### Starting VLC Media Player in Windows

In Windows XP: Click **Start** -\> **Programs** -\> **VideoLAN** -\> **VLC media player**.

In Windows 7: Click **Start** -\> **All Programs** -\> **VideoLAN** -\> **VLC media player**.

VLC is shown on the screen and a small icon  is shown in the system tray.

#### Stopping VLC Media Player

There are three ways to quit VLC:

-   Right click the VLC icon () in the tray and select **Quit** (*Alt-F4*).
-   Click the **Close** button in the main interface of the application.
-   In the **Media** menu, select **Quit** (*Ctrl-Q*).

#### Notification Area Icon

Clicking this icon shows or hides the VLC interface. Hiding VLC does not exit the application. VLC keeps running in the background when it is hidden. Right clicking the icon in the notification area shows a menu with basic operations, such as opening, playing, stopping, or changing a media file.

#### Main Interface

The main interface has the following areas:

-   **Menu bar**.
-   **Track slider** - The track slider is below the menu bar. It shows the playing progress of the media file. You can drag the track slider left to rewind or right to forward the track being played. When a video file is played, the video is shown between the menu bar and the track slider.
    **Note: When a media file is streamed, the track slider does not move because VLC cannot know the total duration.**
-   **Control Buttons** - The buttons below the track slider cover all the basic playback features.

Click here to view an explanation of every menu item.

#### Opening media

See Documentation:Play HowTo/Basic Use 0.9/Opening modes

#### Streaming Media Files

Streaming is a method of delivering audio or video content across a network without the need to download the media file before it is played. You can view or listen to the content as it arrives. It has the advantage that you don't need to wait for large media files to finish downloading before playing them.

VideoLan is designed to stream MPEG videos on high bandwidth networks. VLC can be used as a server to stream MPEG-1, MPEG-2 and MPEG-4 files, DVDs and live videos on the network in unicast or multicast. Unicast is a process where media files are sent to a single system through the network. Multicast is a process where media files are sent to multiple systems through the network.

VLC is also used as a client to receive, decode and display MPEG streams. MPEG-1, MPEG-2 and MPEG-4 streams received from the network or an external device can be sent to one machine or a group of machines.

**To stream a file**:

1.  From the **Media** menu, select **Open Network Stream**. The *Open Media* dialog box loads with the *Network* tab selected.
2.  In the **Please enter a network URL** text box, Type the network URL.
3.  Click **Play**.

Note: When VLC plays a stream, the track slider shows the progress of the playback.

For more information, refer to Documentation:Streaming HowTo/Receive and Save a Stream

#### Converting and Saving a Media File Format

VLC can convert media files from one format to another.

**To convert a media file**:

1.  From the **Media** menu, select **Convert/Save**. The *Open media* dialog window appears.
2.  Click **Add...**. A file selection dialog window appears.
3.  Select the file you want to convert and click **Open**. The *Convert* dialog window appears.
4.  In the **Destination file** text box, indicate the path and file name where you want to store the converted file.
5.  From the **Profile** drop-down, select a conversion profile.
6.  Click **Start**.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Open Media {#open-media}

#### Play a file

To play a file, open the *Media* menu, and select the *Open File* menu item.

An *Open File* dialog box will appear. Select the file you want to open and select *Open*.

VLC will then start playing the designated file. An alternative is to simply drag 'n' drop your file into the VLC main interface or the playlist window from the file explorer (Finder on Mac OS X).

#### Play a CD/DVD/VCD

To play a CD, VCD or a DVD, open the *Media* menu and select *Open Disc* menu item. In the *Open Disk* dialog box, select the type of media (DVD, SVCD/VCD or Audio CD).

You can either select the drive in which the media is located by selecting the drive letter from the *Disc Device* drop-down list, or you can select the *Browse* button, which will open a dialog box that you can use to browse for the media you wish to play.

If you want to start the DVD or VCD playback from a given title and chapter instead of from the beginning, you can set it using the *Title* and *Chapter* selectors. You can also set the *Audio* and *Subtitles* track using the selectors. There is also an option for *No DVD menus*, when reading a DVD.

To start playback select the *Ok* button.

#### Play a network stream (WebRadio, WebTV, etc.)

To open a network stream, open the *Media* menu and select the *Open Network Stream* menu item.

A dialog box will then open with three user input boxes. The first one is for the user to select the *Protocol* of the stream that they wish to open (HTTP/HTTPS/MMS/FTP/RTSP/RTP/UDP/RDMP). The second box is for the user to input the *Address* of the stream and the third one is for the user to select the appropriate port. However in the latest version of VLC (1.1.5), the user only needs to input the *Address* (examples are shown in image above).

To begin playback, select the *Play* button.

If you get some stuttering during playback, you can try to increase the size of the read buffer. This can be done in the *Open Network Stream* dialog box, by firstly checking the *Show more options* check box then adjusting the *Caching* selector, which allows you to choose the amount of time (in milliseconds) VLC should store data in its buffer before starting playback.

#### Play from an acquisition card

To play from an acquisition open the *File* menu, and select *Open Capture Device*.

From here you can choose the *Capture Mode* and the *Video/Audio Device Name*. The user can also adjust the configuration for these devices by clicking *Configure*. The user is also able to set the size of the video that will be played by the Direct Show plugin and options such as 'Device Properties' and 'Tuner Properties' by clicking *Advanced Options*.

For Video4Linux devices, you can set the name of the video and audio devices using the "Video device name" and "Audio device name" text inputs. The "Advanced options..." button allows you to select some further settings useful in some rare cases, such as the chroma of the input (the way colors are encoded) and the size of the input buffer.

To use a Hauppauge PVR card, select the PVR tab in the "Open" dialog box. Use the "Device" text input to set the device of the card you want to use. You can set the Norm of the tuner (*PAL, SECAM or NTSC*) by using the "Norm" Drop Down. The Frequency selector allows you to set the frequency of the tuner (in kHz), the bitrate selector to set the bitrate of the resulting encoded stream (in bit/s). The "Advanced Options button allows to set some more settings, such as the size of the encoded video (in pixels), its framerate (in frame per second), the interval between 2 key frames, etc.

To start playback from an acquisition card, click *Play*.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Playback {#playback}

VLC media player helps you to create media files. After creating media files, the quality has to be tested. You can test the quality and several other parameters using playback. In playback, you can specify parameters such as time, bookmarks, and titles.

#### Bookmarks

You can mark and locate particular places in an audio or video file using the Bookmarks feature of VLC. If you want to view a particular scene in a movie or listen to certain tune in a song repeatedly, you can create bookmarks.

To bookmark a scene in a video:

1.  From the *Playback* menu select the *Bookmarks* option, and the *Manage Bookmarks*. The *Edit Bookmarks* dialog box opens.
2.  Click *Create* to create a bookmark for the current track. The created bookmark appears in the *Edit Bookmarks* dialog box.
3.  To view a scene that is bookmarked, select a bookmark from *Bookmarks* in the *Playback* menu.

Edit Bookmarks dialog box under Windows in VLC 1.1.5

#### Title

In a DVD format, each movie is referred to by its title or name. A title is displayed whenever a movie is played by any media player. You can view all titles in a folder in a sequential manner.

1.  To open a folder, select *Open Folder* from the *Media* menu. Locate the folder in which the video files are present and click *OK*.
2.  To select a title, click *Title* in the *Playback* menu. The selected title is then played.

#### Chapter

A video can also be divided into chapters. Different chapters can be accessed at random in a video which is being played. Using this option, you can directly view your favourite chapter without having to see the complete video.

To play a chapter:

1.  Select *Open Folder* from the *Media* menu.
2.  Locate the folder in which the video files are present.
3.  Select a video file and click *OK*.
    The file is played in the VLC media player.
4.  Select *Chapter* in the *Playback* menu to view the list of chapters. Select a chapter of your choice.

Then selected chapter is played.

#### Navigation

In VLC, you can navigate to different titles and their corresponding chapters. You can also customise a DVD by selecting options such as subtitle, angle and so on.

1.  To customize a title, select the required option from *DVD Menu* in the *Navigation* menu.
2.  To view a title, select a *Title* under *Navigation* in the *Playback* menu. The selected title is played.
3.  To view a chapter in a title, select *Title*. When you select a title, the chapters in a title are listed. Select a chapter.

Refer to #Title and #Chapter sections for more details.

#### Program

This option is enabled only if streams of format DVB and TS are played. Choose the program to select by giving its Service ID. Only use this option if you want to read a multi-program stream (like DVB streams for example). *FIXME: Description needs to be improved*

#### Specify the time

This option is used to go to a specific frame in a media file and listen or view once again.

1.  To specify time select *Jump to Specific Time* from the *Playback* menu. The *Go to Time* dialog box is displayed.
2.  Enter the time in *hh:mm:ss*.
3.  Click on the *Go* button. The control moves the tracker to a specific frame and the media file continues from that specified frame.
4.  Click *Cancel* to exit the dialog box.

Note: Ensure that time limit is within the range of length of the media file.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Playlist {#playlist}

A playlist is a customised list of media files you might want to watch or listen to. Using a playlist, you can specify the media files you want to listen each time you start the VLC media player. You can add tracks from CDs, radio stations, and movies to a playlist. To access the playlist, click on the *Playlist* button in the main interface.

The default playlist view.

#### Additional Sources

In addition to audio and video files, you can play other formats. The additional formats supported by VLC media player are described in the following sections:

-   **Podcasts** - Podcast (Personal On Demand broadCASTING) is a series of audio or video digital media files which is distributed over the Internet and downloaded to media players. If consumers subscribe to Podcasts, whenever new content is added the content gets automatically added to the playlist. You can customise Podcasts. To add a Podcast URL

1.  Select the  *Playlist* button.
2.  Click on the *Internet* button to select it in the left pane. The *Podcasts* menu item will appear under *Internet*.
3.  Select a Podcast stream in the main dialog box. Then right-click the stream and select *Play* from the popup menu.

-   **SAP Announcements** – Helps to advertise your stream over the network.

To play a SAP announcement:

1.  Select the  *Playlist* button.
2.  Click on the *Local Area Network* to select it in the left pane. The *Network Streams (SAP)* menu item will appear under *Local Area Network*.
3.  Select an SAP announcement and right-click. Select *Play* from the popup menu.

-   **Shoutcast Radio Listings** – Shoutcast is a server for streaming the media developed by Nullsoft. Digital audio content can be broadcast from and to media players, and this helps individuals to create Internet radio networks. Using VLC media player, you can listen to your favourite radio stations and you can also create bookmarks to listen to these radio stations in future.

To customize a Shoutcast radio listing:

1.  Select the  *Playlist* button.
2.  Select *Icecast Directory* under *Internet* in the *Playlist* menu. A list of radio stations appears in the right hand panel. If nothing appears in the right hand panel try double-clicking the *Shoutcast Radio* option and wait. It may take a few minutes the first time. After a while, the right hand panel displays a list of titles.

1.  Scroll down and select a radio station.
2.  Right-click on a radio station and:
    1.  Select *Play* if you want to listen to the radio station.
    2.  Select *Remove Selected* if you want to delete the radio station.
    3.  Select the *Stream* option. The Stream output dialog box is displayed. Refer to the Specifying Specifying Streaming options section for more details. Modify the required parameters and click on the *Stream* button to stream the media file.
    4.  Click to select a title in the Playlist dialog box and right-click. Select *Save* from the popup menu. The Stream Output dialog box is displayed. Select the required options and click on the *Save* button in the Stream Output dialog box. Refer to the Specifying Streaming options section for more details.
    5.  Select the Information option. The Media Information dialog box is displayed with details of the media being played.
    6.  Select *Title* to alphabetically sort the radio stations.
    7.  Click to select a title in the Playlist dialog box and right-click. Select *Open Folder* from the popup menu. A folder is opened to show all sub nodes within a title.
    8.  Select *Add Node* to add a node.
    9.  Click to select a title in the Playlist dialog box and right-click. Select *Information* from the popup menu to view the details of the selected title. Refer to the *Media Information* section for more details on options.

-   *Shoutcast TV stream* – You can watch streaming TV using the VLC media player. Shoutcast TV stream refers to a stream transmitted by Nullsoft. The procedure of customising the TV stream and the options are similar to that of the Shoutcast Radio.
-   *Freebox TV listing* – Refers to television service over ADSL accessible by Freebox Free Zone unbundled.

*Note:* You should be connected to the Internet to access these streams.

#### Add Media Files to Playlist

You can add several media files to a playlist. The media files could be selected from the media library, additional sources, and some other source.

To add files to a playlist:

1.  Select the  *Playlist* button.
2.  Right-click on the dialog box and click and a short list appears with two options: Add file and Add directory.
    1.  Select *Add file* to add a file to the playlist.
    2.  Select *Add directory* to add a directory containing media files to the playlist.
3.  Click on the  **Random** icon. This icon toggles between *Random* and *Random Off*. Click on  to play files at random. Click on  and the files are played in an order.
4.  Click on the  *Repeat* icon. This icon toggles between *Repeat One* and *Repeat All*. If you want to listen to a track several times, click on  icon. If you want to listen to all tracks, click on  again.
5.  To search for a media file, enter the name in the *Search* box. To search for media files with certain names or formats, enter a word or phrase in the *Search* box. All files with the specified name are listed.
6.  Click on the  icon. This icon is used to skip to the current item when you have a very long list.
7.  Click the *Remove Selected* button to clear a track from the playlist.

#### Load Playlist

This option is used to add a playlist created in some other media player. You can load playlists of the *.xspf, .asx, .b4s* and *.m3u* formats. To load a playlist:

1.  Select the *Open* option from the *Media* menu. The *Open file* dialog box is displayed.
2.  In the bottom right, change the format to *Playlist Files* in the selector.
3.  Locate a playlist file and click on *Open*.

The selected playlist is added in the current playlist dialog box.

#### Save Playlist

You can save playlists using the VLC media player in format of your choice. To save a playlist:

1.  Create a playlist. Refer to Add Media Files to Playlist for creating a playlist.
2.  Select *Save Playlist to File* from the *Media* menu. The *Choose a filename to save playlist* dialog box is displayed.
3.  Select a name for the playlist.
4.  Select a format in which the playlist must be saved from the *Files of type* list. The Files of type list contains the *.xspf* and *.m3u* formats.
5.  Click on *Save* to save the playlist in the selected format.

#### Play a file

To play a file, open the Media menu, and select the Open File menu item. An Open File dialog box will appear. Select the file you want to open, and click Open. VLC will start playing the selected file. An alternative is to drag 'and' drop your file onto the VLC main interface or playlist window from the file explorer (Finder on MacOS X).

VLC 0.9.8a version Windows XP mode

The File menu - MacOS X interface**- needs verifying for 0.9**

The Open file dialog - wxWidgets interface

(FIXME need 0.9 screenshot for MacOS) The Open file dialog - MacOS X interface

#### Naming Files

You can change the original file name to one you would like before adding the file to VLC. When adding files from the menu bar, the new file name will show in the playlist. However, when dropping the file using the "add/drop" option, VLC may not recognize the name change depending on the file type. In that case, you can right click the header of the playlist column and select "URL," you will then see the original file path for the file.

#### Sorting

In the wxWidgets interface, *Sort* allows you to sort the playlist according to several criteria, or to shuffle it. You can also sort by clicking the header of the column.

In the MacOS X interface, sorting can be done by clicking the header of the column matching the criteria you want to use for sorting.

#### Playlist modes

The playlist supports several playback modes.

In the wxWidgets interface, the toolbar contains three playlist mode buttons. They allow you to enable random mode, to repeat the whole playlist or to repeat one item.

In the MacOS X interface, random mode can be enabled by selecting the *Random* box. A drop down menu allows you to enable playlist and item repeat modes.

#### Misc

##### Search

You also have a search tool. Enter a search string and hit search. The next item to match the string will be highlighted. Keep hitting Search to cycle between all matching items.

##### Moving items

In the wxWidgets interface, the *Up* and *Down* buttons at the bottom of the playlist window allow you to move an item. Select an item and use these buttons to move it.

In the MacOS X interface, you can easily move an item with the mouse, using drag-and-drop.

##### Contextual menu

By right-clicking or control-clicking an item, a contextual menu will appear, giving access to a number of functions (for example, play the item, disable it, delete it, or get info on it).

##### Example finding a Shoutcast radio stream

This example was verified as working on 15 October 2008, using VLC 0.9.4 under Windows Vista. *This needs reproducing by other people on other versions and other operating systems.*

1\. Ensure your firewall is set to allow the VideoLan program to make outgoing connections.

2\. Click *Tools* then *Preferences*, click Interface and then click All under "Show settings". Then click the "-" next to "Playlist" in order to show the "Services discovery" submenu. If the shoutcast radio listings box is empty, click it so that a check-mark appears. The text field underneath should now show the word "shout". Click the Save button to save and close the Preferences window:

3\. Restart VLC media player to make it take notice of the changed preferences.

4\. On the VLC interface click *Playlist*, then click *Show Playlist*. Select the "Shoutcast Radio" in the left hand panel. If nothing appears in the righthand panel, try double-clicking "Shoutcast Radio" and waiting, it may take a few minutes the first time. After a while the righthand panel displays a long list of titles.

5\. Scroll down the radio stations in the right-hand panel and select one. Click the mouse right button and click the "Play" item.

6\. It may take some time for the connection to the radio station to establish (and it may fail if the station's outgoing streams are all occupied). When it does connect, VLC should start playing the audio stream from the station:

##### Example playing a known Shoutcast radio stream

Go to 0 and search for a radio station of your choice. On Windows, right-click your mouse over Shoutcast's "Tunein" button and click "Save Link As..." to save the playlist on your computer. Remember where you saved the playlist, rename it to something that makes sense.

At any time later, you can use VLC to open the saved playlist and listen to that radio station.

For example, to find a BBC World Service radio stream, use a browser to go to: 0

One of the stations listed may be playing the World Service, if so move your mouse over the "TUNEIN!" webicon and click the right mouse button and click "Save Link As...", as described above.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### Snapshots {#snapshots}

There are two ways to take snapshots (i.e., screenshots or frame grabs) with VLC:

1.  Open the *Video* menu, and select the *Take Snapshot* menu item.
2.  Press the snapshot hotkey
    -   Linux / Unix / Windows (Qt Interface): Shift+s
    -   Mac OS X: Command+Alt+s

When a snapshot is taken, it will briefly preview as a thumbnail with its filename and then fade away.

To change the hotkey, go to Tools → Preferences. If "Show settings" is set to Simple, click Hotkeys; if "Show settings" is set to All, navigate to Interface → Hotkeys settings. Set the hotkey for Take video snapshot.

#### Snapshot location, format and name

The snapshot location depends upon your operating system:

-   Windows XP: "%HOMEPATH%\\My Pictures\\"
-   Windows Vista, 7, 8, and 10: "%HOMEPATH%\\Pictures\\"
-   Linux / Unix: \~/Pictures
-   macOS: Desktop/

##### Configuring snapshot options under Windows:

The location, format and name of snapshots may be changed in the *Preferences* menu item in the *Tools* tab, subsection *Video*.

The default format for snapshots is PNG, but this may be changed to JPEG. Also, the default name for snapshots is *vlcsnap-* followed by a timestamp that is *not* the time of the frame in the video you're viewing, but rather the current date and time—as in 2014-01-16-14h57m19s163.

Also, you may substitute other text for *vlcsnap-* in the *Video snapshot file prefix* and you may choose to have snapshots numbered sequentially (i.e., 000001, 000002, 000003, and so on) instead of with a timestamp.

As of version 0.9.0, you may even use variables in the text used for the filename. For example, *\$T* (must be upper case) will insert the video's time code into the file name. If you were to change the prefix to *Friends-\$T-* while watching a DVD of *Friends*, then the snapshot filenames would look something like this: *Friends-00_05_21-2014-01-16-14h57m19s163.png*. This indicates a snapshot taken at 5 minutes and 21 seconds into the video; and it was taken on this day at this time: *2014-01-16-14h57m19s163*.

For a shorter file name, check the "Sequential numbering" option in the configuration box (below). Instead of numbers like *2014-01-16-14h57m19s163*, VLC will simply insert the count of snapshots for that session—for example, *00004*. Thus, in the example above, a snapshot with sequential numbering would look like this: *Friends-00_05_21-000001.png*

For a full list of variables, please see Documentation:Play HowTo/Format String.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

### WebPlugin {#webplugin}

This page is outdated and information might be incorrect.

This Documentation speaks about the VLC media player Web plugins and how to write pages for it.

#### Introduction: Building Web pages with Video

The VLC media player webplugins are native browser plugins, similar to Flash or Silverlight plugins and allow playback inside the browser of all the videos that VLC media player can read.

Additionally to viewing video on all pages, you can build custom pages that will use the advanced features of the plugin, using JavaScript functions to control playback or extract information from the plugin.

There are 2 main plugins: one is ActiveX for IE, the other is NPAPI for the other browsers. They feature the same amount of features.

In older versions, those plugins were very crashy. **We URGE YOU** to use VLC **2.0.0** or newer versions.

##### Browsers support

It has been tested with:

Mozilla Firefox (up to v51)

Internet Explorer

Safari (up to v11)

Chrome (up to v44)

Konqueror

Opera (up to v36)

It has been tested on GNU/Linux, Windows and MacOS.

**Note:**

In the most of browsers, the support for NPAPI plugins was dropped. Only in some forks of Firefox like Waterfox or Pale Moon, NPAPI plugins are still supported.

For this reason, the NPAPI plugin will be dropped in vlc version 4.

#### Embed tag attributes

To embed the plugin into a webpage, use the following 0 template:


If you are using vlc version \< 2.2.0 with Internet Explorer, use instead the following 0 template:


For the declaration of tag attributes, use the tag 0. Here an example:


For compatibility with the mozilla plugin, you can combine both tags:


##### Required elements

These are **required** attributes for the 0 tag:

-   **width**: Specifies the width of the plugin.
-   **height**: Specifies the height of the plugin.
-   **target** (or one of these alias: **mrl**, **filename**, **src**): Specifies the source location (URL) of the video to load.

##### Optional elements

These are additional attributes for the 0 tag:

-   **autoplay**, **autostart**: Specifies whether the plugin starts playing on load. Default: *true*
-   **allowfullscreen** (or **fullscreenEnabled**, **fullscreen**): (since VLC version 2.0.0) Specifies whether the user can switch into fullscreen mode. Default: *true*
-   **windowless**: (since VLC version 2.0.6, only for Mozilla) Draw the video on a window-less (non-accelerated) surface and allow styling (CSS overlay, 3D transformations, and much more). Default: *false*
-   **mute**: Specifies whether the audio volume is initially muted. Default: *false*
-   **volume**: (since VLC version 2.2.2) Specifies the initial audio volume as a percentage. Default: *100*
-   **loop**, **autoloop**: Specifies whether the video loops on end. Default: *false*
-   **controls** (or **toolbar**): Specifies whether the controls are shown by default. Default: *true*
-   **bgcolor**: Specifies the background color of the video player. Default: *#000000*
-   **text**: (only for Mozilla on MacOS) Specifies a text displayed as long as no video is shown. Default: empty
-   **branding**: (in vlc version \< 2.2.2 only for Mozilla on MacOS) Specifies whether VLC branding should be displayed in the web plugin's drawing context. Default: *true*

##### Normal DOM elements

-   **id**: DOM id
-   **name**: DOM name

#### Javascript API description

The vlc plugin exports several objects that can be accessed for setting and getting information. When used improperly the API's will throw an exception that includes a string that explains what happened. For example when you set vlc.audio.track out of range.

##### VLC objects

The vlc plugin knows the following objects:

-   **audio**: Access audio properties.
-   **input**: Access input properties.
    -   **input.title**: Access title properties (available in vlc version ≥ 2.2.2, supported only ≥ 3.0.0)
    -   **input.chapter**: Access chapter properties (available in vlc version ≥ 2.2.2, supported only ≥ 3.0.0)
-   **playlist**: Access playlist properties.
    -   **playlist.items**: Access playlist items properties.
-   **subtitle**: Access subtitle properties.
-   **video**: Access video properties.
    -   **video.deinterlace**: Access deinterlace properties.
    -   **video.marquee**: Access marquee video filter properties.
    -   **video.logo**: Access logo video filter properties.
-   **mediaDescription**: Access media info properties (available in vlc version ≥ 2.0.2).

The following are deprecated:

-   **log**: Access log properties (only available in vlc version ≤ 1.0.0-rc1).
-   **messages**: Access to log message properties (only available in vlc version ≤ 1.0.0-rc1).
-   **iterator**: Access to log iterator properties (only available in vlc version ≤ 1.0.0-rc1).
-   **message**: Access to log message properties (only available in vlc version ≤ 1.0.0-rc1).

###### Example

The following JavaScript code shows howto get a reference to the vlc plugin. This reference can then be used to access the objects of the vlc plugin.

    <!DOCTYPE html>

    VLC Mozilla plugin test page

    <embed type="application/x-vlc-plugin"
           width="640"
           height="480"
           id="vlc" />


    <!--
    var vlc = document.getElementById("vlc");
    vlc.audio.toggleMute();
    //-->


##### Root object

readonly properties

-   **vlc.VersionInfo**: returns version information string

read/write properties

-   *none*

methods

-   **vlc.versionInfo()**: (only for Mozilla) returns version information string (same as VersionInfo)
-   **vlc.getVersionInfo()**: (supported in vlc version ≥ 2.2.2) returns version information string (same as VersionInfo and versionInfo())

&nbsp;

-   **vlc.addEventListener(eventname, callback, bubble)**: (only for Mozilla) add a listener for mentioned event name, callback expects a function and bubble influences the order of eventhandling by JS (usually it is set to false).
-   **vlc.removeEventListener(eventname, callback, bubble)**: (only for Mozilla) remove listener for mentioned event name, callback expects a function and bubble influences the order of eventhandling by JS (usually it is set to false).

&nbsp;

-   **vlc.attachEvent(eventname, callback)**: (only for ActiveX) add listener for mentioned event name, callback expects a function
-   **vlc.detachEvent(eventname, callback)**: (only for ActiveX) remove listener for mentioned event name, callback expects a function

events

-   **MediaPlayerNothingSpecial**: vlc is in idle state doing nothing but waiting for a command to be issued
-   **MediaPlayerOpening**: vlc is opening an media resource locator (MRL)
-   **MediaPlayerBuffering(int cache)**: vlc is buffering
-   **MediaPlayerPlaying**: vlc is playing a media
-   **MediaPlayerPaused**: vlc is in paused state
-   **MediaPlayerStopped**: vlc is in stopped state
-   **MediaPlayerStopAsyncDone**: (supported in vlc version ≥ 3.0.0) playback has stopped asynchronously
-   **MediaPlayerForward**: vlc is fastforwarding through the media (this never gets invoked)
-   **MediaPlayerBackward**: vlc is going backwards through the media (this never gets invoked)
-   **MediaPlayerEncounteredError**: vlc has encountered an error and is unable to continue
-   **MediaPlayerEndReached**: vlc has reached the end of current playlist
-   **MediaPlayerTimeChanged(int time)**: time has changed
-   **MediaPlayerPositionChanged(float position)**: media position has changed
-   **MediaPlayerSeekableChanged(bool seekable)**: media seekable flag has changed (true means media is seekable, false means it is not)
-   **MediaPlayerPausableChanged(bool pausable)**: media pausable flag has changed (true means media is pauseable, false means it is not)
-   **MediaPlayerMediaChanged**: (supported in vlc version ≥ 2.2.0) media has changed
-   **MediaPlayerTitleChanged(int title)**: (in vlc version \< 2.2.0 only for Mozilla) title has changed (DVD/Blu-ray)
-   **MediaPlayerChapterChanged(int chapter)**: (supported in vlc version ≥ 3.0.0) chapter has changed (DVD/Blu-ray)
-   **MediaPlayerLengthChanged(int length)**: (in vlc version \< 2.2.0 only for Mozilla) length has changed
-   **MediaPlayerVout(int count)**: (supported in vlc version ≥ 2.2.7) the number of video output has changed
-   **MediaPlayerMuted**: (supported in vlc version ≥ 2.2.7) audio volume was muted
-   **MediaPlayerUnmuted**: (supported in vlc version ≥ 2.2.7) audio volume was unmuted
-   **MediaPlayerAudioVolume(float volume)**: (supported in vlc version ≥ 2.2.7) audio volume has changed

###### Example

The following code snippet provides easy functions to register and unregister event callbacks on all supported platforms.


    <!--
    function registerVLCEvent(event, handler) {
        var vlc = getVLC("vlc");
        if (vlc) {
            if (vlc.attachEvent) {
                // Microsoft
                vlc.attachEvent (event, handler);
            } else if (vlc.addEventListener) {
                // Mozilla: DOM level 2
                vlc.addEventListener (event, handler, false);
            }
        }
    }
    // stop listening to event
    function unregisterVLCEvent(event, handler) {
        var vlc = getVLC("vlc");
        if (vlc) {
            if (vlc.detachEvent) {
                // Microsoft
                vlc.detachEvent (event, handler);
            } else if (vlc.removeEventListener) {
                // Mozilla: DOM level 2
                vlc.removeEventListener (event, handler, false);
            }
        }
    }
    // event callbacks
    function handle_MediaPlayerNothingSpecial(){
        console.log("Idle");
    }
    function handle_MediaPlayerOpening(){
        console.log("Opening");
    }
    function handle_MediaPlayerBuffering(val){
        console.log("Buffering: " + val + "%");
    }
    function handle_MediaPlayerPlaying(){
        console.log("Playing");
    }
    function handle_MediaPlayerPaused(){
        console.log("Paused");
    }
    function handle_MediaPlayerStopped(){
        console.log("Stopped");
    }
    function handle_MediaPlayerStopAsyncDone(){
        console.log("Stopped asynchronously");
    }
    function handle_MediaPlayerForward(){
        console.log("Forward");
    }
    function handle_MediaPlayerBackward(){
        console.log("Backward");
    }
    function handle_MediaPlayerEndReached(){
        console.log("EndReached");
    }
    function handle_MediaPlayerEncounteredError(){
        console.log("EncounteredError");
    }
    function handle_MediaPlayerTimeChanged(time){
        console.log("Time changed: " + time + " ms");
    }
    function handle_MediaPlayerPositionChanged(val){
        console.log("Position changed: " + val);
    }
    function handle_MediaPlayerSeekableChanged(val){
        console.log("Seekable changed: " + val);
    }
    function handle_MediaPlayerPausableChanged(val){
        console.log("Pausable changed: " + val);
    }
    function handle_MediaPlayerMediaChanged(){
        console.log("Media changed");
    }
    function handle_MediaPlayerTitleChanged(val){
        console.log("Title changed: " + val);
    }
    function handle_MediaPlayerChapterChanged(val){
        console.log("Chapter changed: " + val);
    }
    function handle_MediaPlayerLengthChanged(val){
        console.log("Length changed: " + val + " ms");
    }
    function handle_MediaPlayerVout(val){
        console.log("Number of video output changed: " + val);
    }
    function handle_MediaPlayerMuted(){
        console.log("Audio volume muted");
    }
    function handle_MediaPlayerUnmuted(){
        console.log("Audio volume unmuted");
    }
    function handle_MediaPlayerAudioVolume(volume){
        console.log("Audio volume changed: " + Math.round(volume * 100) + "%");
    }
    // Register a bunch of callbacks.
    registerVLCEvent("MediaPlayerNothingSpecial", handle_MediaPlayerNothingSpecial);
    registerVLCEvent("MediaPlayerOpening", handle_MediaPlayerOpening);
    registerVLCEvent("MediaPlayerBuffering", handle_MediaPlayerBuffering);
    registerVLCEvent("MediaPlayerPlaying", handle_MediaPlayerPlaying);
    registerVLCEvent("MediaPlayerPaused", handle_MediaPlayerPaused);
    registerVLCEvent("MediaPlayerStopped", handle_MediaPlayerStopped);
    registerVLCEvent("MediaPlayerStopAsyncDone", handle_MediaPlayerStopAsyncDone);
    registerVLCEvent("MediaPlayerForward", handle_MediaPlayerForward);
    registerVLCEvent("MediaPlayerBackward", handle_MediaPlayerBackward);
    registerVLCEvent("MediaPlayerEndReached", handle_MediaPlayerEndReached);
    registerVLCEvent("MediaPlayerEncounteredError", handle_MediaPlayerEncounteredError);
    registerVLCEvent("MediaPlayerTimeChanged", handle_MediaPlayerTimeChanged);
    registerVLCEvent("MediaPlayerPositionChanged", handle_MediaPlayerPositionChanged);
    registerVLCEvent("MediaPlayerSeekableChanged", handle_MediaPlayerSeekableChanged);
    registerVLCEvent("MediaPlayerPausableChanged", handle_MediaPlayerPausableChanged);
    registerVLCEvent("MediaPlayerMediaChanged", handle_MediaPlayerMediaChanged);
    registerVLCEvent("MediaPlayerTitleChanged", handle_MediaPlayerTitleChanged);
    registerVLCEvent("MediaPlayerChapterChanged", handle_MediaPlayerChapterChanged);
    registerVLCEvent("MediaPlayerLengthChanged", handle_MediaPlayerLengthChanged);
    registerVLCEvent("MediaPlayerVout", handle_MediaPlayerVout);
    registerVLCEvent("MediaPlayerMuted", handle_MediaPlayerMuted);
    registerVLCEvent("MediaPlayerUnmuted", handle_MediaPlayerUnmuted);
    registerVLCEvent("MediaPlayerAudioVolume", handle_MediaPlayerAudioVolume);
    //-->


###### Event registration issue with IE11

Since IE11, the methods attachEvent() and detachEvent() are not longer available. So the registration of events is not possible. But there are two workarounds:

Workaround 1 - Use the IE10 compatibility mode to re-enable the missing methods. For using the compatibility mode, add this meta tag to the header of your page:


Workaround 2 - Use an old IE-only implementation of event registration. An example with the MediaPlayerBuffering event:


    console.log("Buffering: " + val + "%");


##### Audio object

readonly properties

-   **vlc.audio.count**: (supported in vlc version ≥ 1.1.0) returns the number of audio track available.

read/write properties

-   **vlc.audio.mute**: boolean value to mute and unmute the audio.
-   **vlc.audio.volume**: a value between \[0-200\] which indicates a percentage of the volume.
-   **vlc.audio.track**: (supported in vlc version \> 0.8.6) a value between \[1-65535\] which indicates the audio track to play or that is playing. a value of 0 means the audio is/will be disabled.
-   **vlc.audio.channel**: (supported in vlc version \> 0.8.6) integer value between \[1-5\] that indicates which audio channel mode is used, values can be: "1=stereo", "2=reverse stereo", "3=left", "4=right", "5=dolby". Use vlc.audio.channel to check if setting of the audio channel mode has succeeded.

methods

-   **vlc.audio.toggleMute()**: boolean toggle that mutes and unmutes the audio based upon the previous state.
-   **vlc.audio.description(int i)**: (supported in vlc version ≥ 1.1.0) give the i-th audio track name. 0 corresponds to disable and 1 to the first audio track.

###### Example

    Audio Channel:

        Stereo
        Reverse stereo
        Left
        Right
        Dolby


    <!--
    function doAudioChannel(value)
    {
        var vlc = getVLC("vlc");
        vlc.audio.channel = parseInt(value);
        alert(vlc.audio.channel);
    }
    //-->


##### Input object

readonly properties

-   **vlc.input.length**: length of the input file in number of milliseconds. 0 is returned for 'live' streams or clips whose length cannot be determined by VLC. It returns -1 if no input is playing.
-   **vlc.input.fps**: frames per second returned as a float (typically 60.0, 50.0, 23.976, etc...)
-   **vlc.input.hasVout**: a boolean that returns true when the video is being displayed, it returns false when video is not displayed

read/write properties

-   **vlc.input.position**: normalized position in multimedia stream item given as a float value between \[0.0 - 1.0\]
-   **vlc.input.time**: the absolute position in time given in milliseconds, this property can be used to seek through the stream

&nbsp;

     <!-- absolute seek in stream -->
     vlc.input.time =
     <!-- relative seek in stream -->
     vlc.input.time = vlc.input.time +

-   **vlc.input.state**: current state of the input chain given as enumeration:

&nbsp;

-   **0**: IDLE
-   **1**: OPENING
-   **2**: BUFFERING
-   **3**: PLAYING
-   **4**: PAUSED
-   **5**: STOPPING
-   **6**: ENDED
-   **7**: ERROR

Note: Test for ENDED=6 to catch end of playback. Checking for STOPPING=5 is NOT ENOUGH.

-   **vlc.input.rate**: input speed given as float (1.0 for normal speed, 0.5 for half speed, 2.0 for twice as fast, etc.).

&nbsp;

-   **rate \> 1**: fast forward
-   **rate = 1**: normal speed
-   **rate \< 1**: slow motion

methods

-   *none*

###### Title object

readonly properties

-   **vlc.input.title.count**: (supported in vlc version ≥ 2.2.2) returns the number of title available.

read/write properties

-   **vlc.input.title.track**: (supported in vlc version ≥ 2.2.2) get and set the title track. The property takes an integer as input value \[0..65535\]. It returns -1 if no titles are available.

methods

-   **vlc.input.title.description(int i)**: (supported in vlc version ≥ 2.2.2) give the i-th title name.

###### Chapter object

readonly properties

-   **vlc.input.chapter.count**: (supported in vlc version ≥ 2.2.2) returns the number of chapter available in the current title.

read/write properties

-   **vlc.input.chapter.track**: (supported in vlc version ≥ 2.2.2) get and set the chapter track. The property takes an integer as input value \[0..65535\]. It returns -1 if no chapters are available.

methods

-   **vlc.input.chapter.description(int i)**: (supported in vlc version ≥ 2.2.2) give the i-th chapter name.
-   **vlc.input.chapter.countForTitle(int i)**: (supported in vlc version ≥ 2.2.2) returns the number of chapter available for a specific title.
-   **vlc.input.chapter.prev()**: (supported in vlc version ≥ 2.2.2) play the previous chapter.
-   **vlc.input.chapter.next()**: (supported in vlc version ≥ 2.2.2) play the next chapter.

##### Playlist object

readonly properties

-   **vlc.playlist.itemCount**: number that returns the amount of items currently in the playlist (**deprecated**, do not use, see Playlist items)
-   **vlc.playlist.isPlaying**: a boolean that returns true if the current playlist item is playing and false when it is not playing
-   **vlc.playlist.currentItem**: (supported in vlc version ≥ 2.2.0) number that returns the index of the current item in the playlist. It returns -1 if the playlist is empty or no item is active.
-   **vlc.playlist.items**: return the playlist items collection, see Playlist items

read/write properties

-   *none*

methods

-   **vlc.playlist.add(mrl)**: add a playlist item as MRL. The MRL must be given as a string. Returns the index of the just added item in the playlist as a number.
-   **vlc.playlist.add(mrl,name,options)**: add a playlist item as MRL, with metaname 'name' and options 'options'. options are text arguments which can be provided either as a single string containing space separated values, akin to VLC command line, or as an array of string values. Returns the index of the just added item in the playlist as a number.

&nbsp;

    var options = new Array(":aspect-ratio=4:3", "--rtsp-tcp");
    // Or: var options = ":aspect-ratio=4:3 --rtsp-tcp";
    var id = vlc.playlist.add("rtsp://servername/item/to/play", "fancy name", options);
    vlc.playlist.playItem(id);

-   **vlc.playlist.play()**: start playing the current playlist item
-   **vlc.playlist.playItem(number)**: start playing the item whose identifier is number
-   **vlc.playlist.pause()**: pause the current playlist item
-   **vlc.playlist.togglePause()**: toggle the pause state for the current playlist item
-   **vlc.playlist.stop()**: stop playing the current playlist item
-   **vlc.playlist.stop_async()**: (supported in vlc version ≥ 3.0.0) stop playing the current playlist item asynchronously and fire the event MediaPlayerStopAsyncDone, if done
-   **vlc.playlist.next()**: iterate to the next playlist item
-   **vlc.playlist.prev()**: iterate to the previous playlist item
-   **vlc.playlist.clear()**: empty the current playlist, all items will be deleted from the playlist (**deprecated**, do not use, see Playlist items)
-   **vlc.playlist.removeItem(number)**: remove the item from playlist whose identifier is number (**deprecated**, do not use, see Playlist items)
-   **vlc.playlist.parse(options, timeout)**: (supported in vlc version ≥ 3.0.0) Parse the first media in the playlist. This fetches (local or network) art, meta data and/or tracks information. A timeout for parsing can be set in milliseconds or to indefinitely (0). Returns the parsed status.

Available options flags for parsing (which can be combined):

-   **0**: Parse media if it's a local file.
-   **1**: Parse media even if it's a network file.
-   **2**: Fetch meta and covert art using local resources.
-   **4**: Fetch meta and covert art using network resources.
-   **8**: Interact with the user. Set this flag in order to receive a callback when the input is asking for credentials.

Parsed status given as enumeration:

-   **1**: skipped
-   **2**: failed
-   **3**: timeout
-   **4**: done

###### Playlist items object

readonly properties

-   **vlc.playlist.items.count**: number of items currently in the playlist

read/write properties

-   *none*

methods

-   **vlc.playlist.items.clear()**: empty the current playlist, all items will be deleted from the playlist. (note: if a movie is playing, it will not stop)
-   **vlc.playlist.items.remove(number)**: remove the item whose identifier is number from playlist. (note: this number is the current position in the playlist. It's not the number given by vlc.playlist.add(), if any items of the playlist were removed in the meantime.)

##### Subtitle object

readonly properties

-   **vlc.subtitle.count**: (supported in vlc version ≥ 1.1.0) returns the number of subtitle available.

read/write properties

-   **vlc.subtitle.track**: (supported in vlc version ≥ 1.1.0) get and set the subtitle track to show on the video screen. The property takes an integer as input value \[1..65535\]. If subtitle track is set to 0, the subtitles will be disabled.

methods

-   **vlc.subtitle.description(int i)**: (supported in vlc version ≥ 1.1.0) give the i-th subtitle name. 0 correspond to disable and 1 to the first subtitle.

##### Video object

readonly properties

-   **vlc.video.width**: returns the horizontal size of the video
-   **vlc.video.height**: returns the vertical size of the video
-   **vlc.video.count**: (supported in vlc version ≥ 2.2.7) returns the number of video track available.

read/write properties

-   **vlc.video.fullscreen**: when set to true the video will be displayed in fullscreen mode, when set to false the video will be shown inside the video output size. The property takes a boolean as input.
-   **vlc.video.aspectRatio**: get and set the aspect ratio to use in the video screen. The property takes a string as input value. Typical values are: "1:1", "4:3", "16:9", "16:10", "221:100" and "5:4"
-   **vlc.video.scale**: (supported in vlc version ≥ 3.0.0) get and set the video scaling factor as float. That is the ratio of the number of pixels on screen to the number of pixels in the original decoded video in each dimension. Zero is a special value; it will adjust the video to the output window.
-   **vlc.video.subtitle**: (supported in vlc version \> 0.8.6a) get and set the subtitle track to show on the video screen. The property takes an integer as input value \[1..65535\]. If subtitle track is set to 0, the subtitles will be disabled.
-   **vlc.video.crop**: (removed with vlc version 4.0.0) get and set the geometry of the zone to crop. This is set as \ x \ + \ + \. A possible value is: "120x120+10+10"
-   **vlc.video.teletext**: (supported in vlc version ≥ 0.9.0) get and set teletext page to show on the video stream. This will only work if a teletext elementary stream is available in the video stream. The property takes an integer as input value \[0..1000\] for indicating the teletext page to view, setting the value to 0 means hide teletext.
-   **vlc.video.track**: (supported in vlc version ≥ 2.2.7) a value between \[1-65535\] which indicates the video track to play or that is playing. a value of 0 means the video is/will be disabled.

methods

-   **vlc.video.takeSnapshot()**: (supported in vlc version ≥ 0.9.0, only for ActiveX) generates a snapshot and saves it on the desktop
-   **vlc.video.toggleFullscreen()**: toggle the fullscreen mode based on the previous setting
-   **vlc.video.toggleTeletext()**: (supported in vlc version ≥ 0.9.0) toggle the teletext page to overlay transparent or not, based on the previous setting
-   **vlc.video.description(int i)**: (supported in vlc version ≥ 2.2.7) give the i-th video track name. 0 corresponds to disable and 1 to the first video track.
-   **vlc.video.crop_ratio(int numerator, int denominator)**: (supported in vlc version ≥ 4.0.0) Forces a crop ratio on any and all video tracks rendered by the media player. To disable video crop, set a crop ratio with zero as denominator.
-   **vlc.video.crop_window(int x, int y, int width, int height)**: (supported in vlc version ≥ 4.0.0) Selects a sub-rectangle of video to show. Any pixels outside the rectangle will not be shown. To unset the video crop window, use vlc.video.crop_ratio() or vlc.video.crop_border().
-   **vlc.video.crop_border(int left, int right, int top, int bottom)**: (supported in vlc version ≥ 4.0.0) Selects the size of video edges to be cropped out. To unset the video crop borders, set all borders to zero.

###### Deinterlace Object

readonly properties

-   *none*

read/write properties

-   *none*

methods

-   **vlc.video.deinterlace.enable("my_mode")**: (supported in vlc version ≥ 1.1.0) enable deinterlacing with my_mode. You can enable it with "blend", "bob", "discard", "linear", "mean", "x", "yadif" or "yadif2x" mode. Enabling too soon deinterlacing may cause some problems. You have to wait that all variable are available before enabling it.
-   **vlc.video.deinterlace.disable()**: (supported in vlc version ≥ 1.1.0) disable deinterlacing.

###### Marquee Object

readonly properties

-   *none*

read/write properties

-   **vlc.video.marquee.text**: (supported in vlc version ≥ 1.1.0, since vlc version 4.0.0 writeonly) display my text on the screen.
-   **vlc.video.marquee.color**: (supported in vlc version ≥ 1.1.0) change the text color. val is the new color to use (WHITE=0x000000, BLACK=0xFFFFFF, RED=0xFF0000, GREEN=0x00FF00, BLUE=0x0000FF...).
-   **vlc.video.marquee.opacity**: (supported in vlc version ≥ 1.1.0) change the text opacity, val is defined from 0 (completely transparent) to 255 (completely opaque).
-   **vlc.video.marquee.position**: (supported in vlc version ≥ 1.1.0) change the text position ("center", "left", "right", "top", "top-left", "top-right", "bottom", "bottom-left", "bottom-right").
-   **vlc.video.marquee.refresh**: (supported in vlc version ≥ 1.1.0) change the marquee refresh period.
-   **vlc.video.marquee.size**: (supported in vlc version ≥ 1.1.0) val define the new size for the text displayed on the screen. If the text is bigger than the screen then the text is not displayed.
-   **vlc.video.marquee.timeout**: (supported in vlc version ≥ 1.1.0) change the timeout value. val is defined in ms, but 0 value correspond to unlimited.
-   **vlc.video.marquee.x**: (supported in vlc version ≥ 1.1.0) change text abscissa.
-   **vlc.video.marquee.y**: (supported in vlc version ≥ 1.1.0) change text ordinate.

methods

-   **vlc.video.marquee.enable()**: (supported in vlc version ≥ 1.1.0) enable marquee filter.
-   **vlc.video.marquee.disable()**: (supported in vlc version ≥ 1.1.0) disable marquee filter.

Some problems may happen (option like color or text will not be applied) because of the VLC asynchronous functioning. To avoid it, after enabling marquee, you have to wait a little time before changing an option. But it should be fixed by the new vout implementation.

NOTE: see [this forum post](https://forum.videolan.org/viewtopic.php?f=16&t=89427#p295058)

###### Logo Object

readonly properties

-   *none*

read/write properties

-   **vlc.video.logo.opacity**: (supported in vlc version ≥ 1.1.0) change the picture opacity, val is defined from 0 (completely transparent) to 255 (completely opaque).
-   **vlc.video.logo.position**: (supported in vlc version ≥ 1.1.0) change the text position ("center", "left", "right", "top", "top-left", "top-right", "bottom", "bottom-left", "bottom-right").
-   **vlc.video.logo.delay**: (supported in vlc version ≥ 1.1.0) display each picture for a duration of 1000 ms (default) before displaying the next picture.
-   **vlc.video.logo.repeat**: (supported in vlc version ≥ 1.1.0) number of loops for picture animation (-1=continuous, 0=disabled, n=n-times). The default is -1 (continuous).
-   **vlc.video.logo.x**: (supported in vlc version ≥ 1.1.0) change the x-offset for displaying the picture counting from top-left on the screen.
-   **vlc.video.logo.y**: (supported in vlc version ≥ 1.1.0) change the y-offset for displaying the picture counting from top-left on the screen.

methods

-   **vlc.video.logo.enable()**: (supported in vlc version ≥ 1.1.0) enable logo video filter.
-   **vlc.video.logo.disable()**: (supported in vlc version ≥ 1.1.0) disable logo video filter.
-   **vlc.video.logo.file("file.png")**: (supported in vlc version ≥ 1.1.0) display my file.png as logo on the screen.

Some problems may happen because of the VLC asynchronous functioning. To avoid it, after enabling logo video filter, you have to wait a little time before changing an option. But it should be fixed by the new vout implementation.

##### MediaDescription Object

readonly properties

-   **vlc.mediaDescription.title**: (supported in vlc version ≥ 2.0.2) returns title meta information field.
-   **vlc.mediaDescription.artist**: (supported in vlc version ≥ 2.0.2) returns artist meta information field.
-   **vlc.mediaDescription.genre**: (supported in vlc version ≥ 2.0.2) returns genre meta information field.
-   **vlc.mediaDescription.copyright**: (supported in vlc version ≥ 2.0.2) returns copyright meta information field.
-   **vlc.mediaDescription.album**: (supported in vlc version ≥ 2.0.2) returns album meta information field.
-   **vlc.mediaDescription.trackNumber**: (supported in vlc version ≥ 2.0.2) returns trackNumber meta information field.
-   **vlc.mediaDescription.description**: (supported in vlc version ≥ 2.0.2) returns description meta information field.
-   **vlc.mediaDescription.rating**: (supported in vlc version ≥ 2.0.2) returns rating meta information field.
-   **vlc.mediaDescription.date**: (supported in vlc version ≥ 2.0.2) returns date meta information field.
-   **vlc.mediaDescription.setting**: (supported in vlc version ≥ 2.0.2) returns setting meta information field.
-   **vlc.mediaDescription.URL**: (supported in vlc version ≥ 2.0.2) returns URL meta information field.
-   **vlc.mediaDescription.language**: (supported in vlc version ≥ 2.0.2) returns language meta information field.
-   **vlc.mediaDescription.nowPlaying**: (supported in vlc version ≥ 2.0.2) returns nowPlaying meta information field.
-   **vlc.mediaDescription.publisher**: (supported in vlc version ≥ 2.0.2) returns publisher meta information field.
-   **vlc.mediaDescription.encodedBy**: (supported in vlc version ≥ 2.0.2) returns encodedBy meta information field.
-   **vlc.mediaDescription.artworkURL**: (supported in vlc version ≥ 2.0.2) returns artworkURL meta information field.
-   **vlc.mediaDescription.trackID**: (supported in vlc version ≥ 2.0.2) returns trackID meta information field.

read/write properties

-   *none*

methods

-   *none*

##### DEPRECATED APIs

###### DEPRECATED: Log object

**CAUTION**: For security concern, VLC 1.0.0-rc1 is the latest (near-to-stable) version in which this object and its children are supported.

This object allows accessing VLC main message logging queue. Typically this queue capacity is very small (no more than 256 entries) and can easily overflow, therefore messages should be read and cleared as often as possible.

readonly properties

-   **vlc.log.messages**: returns the message collection, see Messages object

read/write properties

-   **vlc.log.verbosity**: write number \[-1,0,1,2,3\] for changing the verbosity level of the log messages; messages whose verbosity is higher than set will be not be logged in the queue. The numbers have the following meaning: -1 disable, 0 info, 1 error, 2 warning, 3 debug.

methods

-   *none*

###### DEPRECATED: Messages object

**CAUTION**: For security concern, VLC 1.0.0-rc1 is the latest (near-to-stable) version in which this object and its children are supported.

readonly properties

-   **messages.count**: returns number of messages in the log

read/write properties

-   *none*

methods

-   **messages.clear()**: clear the current log buffer. It should be called as frequently as possible to not overflow the message queue. Call this method after the log messages of interest are read.
-   **messages.iterator()**: creates and returns an iterator object, used to iterate over the messages in the log. **Don't clear the log buffer while holding an iterator object.**

###### DEPRECATED: Messages Iterator object

**CAUTION**: For security concern, VLC 1.0.0-rc1 is the latest (near-to-stable) version in which this object and its children are supported.

readonly properties

-   **iterator.hasNext**: returns a boolean that indicates whether *vlc.log.messages.next()* will return the next message.

read/write properties

-   *none*

methods

-   **iterator.next()**: returns the next message object in the log, see Message object

###### DEPRECATED: Message subobject

**CAUTION**: For security concern, VLC 1.0.0-rc1 is the latest (near-to-stable) version in which this object and its children are supported.

-   **message.severity**: number that indicates the severity of the log message (0 = info, 1 = error, 2 = warning, 3 = debug)
-   **message.name**: name of VLC module that printed the log message (e.g: main, http, directx, etc...)
-   **message.type**: type of VLC module that printed the log message (eg: input, access, vout, sout, etc...)
-   **message.message**: the message text

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**

## VLC on Phones and Tablets {#vlc-on-phones-and-tablets}

### Android {#android}

------ Work in progress ------

Here is the documentation of the Android port of VLC media player.

#### Preliminary Notes

VLC for Android is a little different from VLC on desktops. In some ways, you can do more; in other ways, you can do less. VLC for Android only does media playback. Active streaming or file / stream to file conversations are not supported for usability and performance reasons. This walk-through does only include screenshots of a phone interface for size reasons. However, all features are also available on tablets with a similar appearance.

#### Feature Overview

-   **Feature**: Version 1.0 Version 1.6 Version 2.0 Version 2.5 Version 3.0 Version 3.1
-   **Opening Network Streams**: No Yes Yes Yes Yes Yes
-   **UPnP discovery and streaming**: No Yes Yes Yes Yes Yes
-   **Plex server discovery and streaming**: No Yes Yes Yes Yes Yes
-   **Password-protected Plex shares**: No No No No No No
-   **Downloads from UPnP multimedia servers**: No No No No No No
-   **FTP discovery, streaming**: No Yes Yes Yes Yes Yes
-   **Store FTP server bookmarks**: No No No Yes Yes Yes
-   **Audio Playback via Connector Cables**: Yes Yes Yes Yes Yes Yes
-   **Video Playback via Connector Cables**: Yes Yes Yes Yes Yes Yes
-   **Subtitles playback**: Yes Yes Yes Yes Yes Yes
-   **Subtitles Font Customization**: No No Yes Yes Yes Yes
-   **Closed Caption playback**: Yes Yes Yes Yes Yes Yes
-   **Teletext subtitles playback**: No No Yes Yes Yes Yes
-   **Multi-track audio handling**: No Yes Yes Yes Yes Yes
-   **Video Filtering incl. Screen Brightness**: No No No No No No
-   **Video Cropping and Aspect Ratio variation**: Yes Yes Yes Yes Yes Yes
-   **Deinterlacing**: Yes Yes Yes Yes Yes Yes
-   **Playback Speed control**: Yes Yes Yes Yes Yes Yes
-   **Audio/Subtitles delay control**: No No Yes Yes Yes Yes
-   **Repeated playback**: Yes Yes Yes Yes Yes Yes
-   **Gestures based playback control**: Yes Yes Yes Yes Yes Yes
-   **Playback of Audio-only media (mp3, m4a, flac, …)**: Yes Yes Yes Yes Yes Yes
-   **Audio Playback in Background**: Yes Yes Yes Yes Yes Yes
-   **Video Playback in Background**: Yes Yes Yes Yes Yes Yes
-   **Playback timer**: Yes Yes Yes Yes Yes Yes
-   **Chapter & title selection**: No No Yes Yes Yes Yes
-   **10-band equalizer**: Yes Yes Yes Yes Yes Yes
-   **Playback UI Lock**: Yes Yes Yes Yes Yes Yes
-   **Smart Media Library sorting for audio albums and TV shows**: Yes Yes Yes Yes Yes Yes
-   **Media Library Search**: No Yes Yes Yes Yes Yes
-   **Passcode Lock**: No No No No No No
-   **Voice search support**: No No No Yes Yes Yes
-   **Voice actions support**: No No No No No No
-   **Organize media in folders**: No No No No No No
-   **Use folders as playlists**: No No Yes Yes Yes Yes
-   **Loop playlists**: Yes Yes Yes Yes Yes Yes
-   **Playback control through headphones or lock screen**: Yes Yes Yes Yes Yes Yes
-   **Mediasession support (Wear, TV, etc…)**: No No Partial Yes Yes Yes
-   **Playback is paused when headphones are unplugged**: Yes Yes Yes Yes Yes Yes
-   **WiFi upload and HTTP downloads in background**: No No No No No No
-   **Support for password protected HTTP streams**: No No No No No No
-   **Sharing files with further apps**: No No No No No No
-   **Custom vlc:// protocol**: No No Yes Yes Yes Yes
-   **Support for x-callback-url**: No No No No No Yes
-   **Action mode**: No No No Yes Yes Yes
-   **Android TV**: No Yes Yes Yes Yes Yes
-   **Picture-in-Picture**: No No Partial Yes Yes Yes
-   **ChromeOS support**: No ARC ARC Yes Yes Yes
-   **Android Auto**: No No No Yes No Yes
-   **Sorting**: No No Partial Yes Yes Yes
-   **360° videos**: No No No Yes Yes Yes
-   **DayNight mode**: No No No Yes Yes Yes
-   **Chromecast**: No No No No Yes Yes
-   **Equalizer custom presets**: No No No No Yes Yes
-   **Audio boost**: No No No No Yes Yes
-   **Android 2.1 support**: Yes Yes No No No No
-   **Android 2.2 support**: Yes Yes Yes No No No
-   **Android 2.3 support**: Yes Yes Yes Yes Yes No
-   **Android 6 (Runtime permissions)**: No No Yes Yes Yes Yes
-   **Android 8 support**: No No No Partial Partial Yes

#### Installation

There are many ways to install VLC on Android. This may be because you have a non-ARMv7 or x86 processor or do not wish to use the Play Store for whatever reason.

##### From the Play Store (recommended)

The normal way, for ARMv7 (and above) and x86 processors only. Don't know your processor? Don't worry, if you can download it, you have a compatible ARMv7 or an x86 processor.

0

##### From the F-Droid Repository

The F-Droid repository (0) is a completely FOSS (Free and Open Source Software) equivalent to the Google Play Store. The F-Droid Repository and all apps within it are provided completely free of charge and licensed under open source licenses. The F-Droid repository can be downloaded directly from their website. The "Unknown Sources" setting must be turned on for Android devices (typically located in Settings -\> Security) in order to install repositories other than the Google Play Store.

##### From VideoLAN

If you can't download from the Play Store or just want to install the VLC .apk by yourself, follow these steps:

1.  Go to Android Settings → Security → Device Administration → Enable 'Unknown Sources'
2.  Go to our download server, preferably from your device: 0
3.  Choose your processor architecture (ARMv7 or Intel x86) and grab the .apk file.
4.  Click on the .apk you just downloaded and install it.

Don't really know your processor architecture? Try both... it's not very clever, but it's harmless.

None of the two work? It is possible that you have an older processor with the ARMv6 architecture. The solution for now is to install a Nightly Build release. See below.

Still doesn't work? Really? Well, then you must have an exotic processor... Contact us, on the [Android forum](http://forum.videolan.org/viewforum.php?f=35) or directly at [android-support@videolan.org](mailto:android-support@videolan.org).

##### Be a Beta tester or try a Beta release

You want want to know the future of VLC for Android? You want to help us and/or test if your issue is already fixed for the next release ?

###### Be a Beta tester

Just follow this link [Be a Beta tester](https://play.google.com/apps/testing/org.videolan.vlc)

Soon, Beta release will automatically install on our device.

###### Try a Beta release

You don't want to be a Beta Tester but just try a Beta ? Follow these steps :

1.  Go to Android Settings → Security → Device Administration → Enable 'Unknown Sources'
2.  Go to our server, preferably from your device, : 0
3.  Choose your processor architecture (ARMv7, ARMv8, x86...)
    Don't really know your processor architecture? Try both... it's not very clever, but it's harmless
4.  Download the chosen .apk on your device
5.  Click on the .apk you just download and install it.

##### Install a Nightly Build

You fear nothing and want our very last works on VLC ? Or you have an ARMv6 Processor and want VLC? Follow these steps:

1.  Go to Android Settings → Security → Device Administration → Enable 'Unknown Sources'
2.  Go to our server, preferably from your device: 0
3.  Choose your processor architecture (ARMv7, ARMv8, x86...)
4.  Grab the latest .apk
5.  Click on the .apk you just download and install it.

You may experience some weird issues but generally, it works fine. If not, please try an older nightly release, and contact us.

#### Interface

At first start, VLC scans all your device to find all your media files. This is the main interface after the scan :

Show Menu

Video view

Audio view

Directory view

History view

Video browser view

Search a specific media

Open network MRL

Load last playlist

More actions :

-   Sort by name or length
-   Refresh your media library
-   Equalizer
-   Preferences
-   About VLC

#### Playing Video

##### Video browser view

This view displays all your videos present in your device, or in the directories you have specified (see Preferences). To play one, just click on it, like the video .

Note the difference with the video  which is a group of videos : VLC automatically groups your videos with the 4 same starting letters.

A Video

A group of videos.

##### Video playback interface

Video title

Battery and time

Play / Pause

Aspect ratio

Audio tracks

Subtitles tracks

Video menu (for DVD iso)

Lock screen

Elapsed time

Seek bar

Total time / Remaining time

Advanced Options

-   Playback Speed
-   Sleep timer
-   Jump to specific time
-   Add subtitle

Some precisions:

-   You can change audio and/or subtitle track if there are any. If not, these icons won't be displayed.
-   The Video Menu icon is only displayed for iso video (a DVD iso for example)

##### Video playback gesture

Adjust Brightness

Adjust Volume

Quick search

#### Playing Audio

TODO

-   You can change the time display to remaining time (e.g. -1:30 for 1:30 minutes remaining) in the audio player by tapping on the current time label in the left.

#### Settings

#### See Also

AndroidFAQ
Android Checklist
Android Player Intents
Android Report bugs

### IOS {#ios}

#### Preliminary Note

VLC for iOS is different from what you can do with VLC on desktops. In some ways, you can do more; in other ways, you can do less. VLC for iOS only does media playback. Active streaming or file / stream-to-file conversations are not supported for usability and performance reasons. This walkthrough only includes screenshots of the iPhone interface for size reasons. However, all features are also available on iPad with virtually the same appearance.

#### Feature Overview

VLC for iOS 2 is a full re-write of the original app and shares no code with it. It is under active development and evolves over time. It is strongly recommended to always use the latest version. To keep track of features added over time, here's a chart:

-   **Feature**: Version 1.x Version 2.0 Version 2.1 Version 2.2 Version 2.3 Version 2.4 Version 2.5 Version 2.6
-   **iTunes File Sharing**: Yes Yes Yes Yes Yes Yes Yes Yes
-   **WiFi Upload**: No Yes Yes Yes Yes Yes Yes Yes
-   **Download from device via WiFi**: No No No No No Yes Yes Yes
-   **Box Integration**: No No No No No No Yes Yes
-   **Streaming from Box**: No No No No No No Yes Yes
-   **Dropbox Integration**: No Yes Yes Yes Yes Yes Yes Yes
-   **Streaming from Dropbox**: No No No Yes Yes Yes Yes Yes
-   **iCloud Integration**: No No No No No No Yes Yes
-   **Streaming from iCloud**: No No No No No No No No
-   **Google Drive integration**: No No No Yes Yes Yes Yes Yes
-   **Streaming from Google Drive**: No No No No No Yes Yes Yes
-   **OneDrive Integration**: No No No No No No Yes Yes
-   **Streaming from OneDrive**: No No No No No No Yes Yes
-   **HTTP Downloads from Web**: No Yes Yes Yes Yes Yes Yes Yes
-   **FTP Downloads from Web**: No No Yes Yes Yes Yes Yes Yes
-   **Opening Network Streams**: No GUI Yes Yes Yes Yes Yes Yes Yes
-   **UPnP discovery and streaming**: No No Yes Yes Yes Yes Yes Yes
-   **Plex server discovery and streaming**: No No No No No Yes Yes Yes
-   **Password-protected Plex shares**: No No No No No No No Yes
-   **Downloads from UPnP multimedia servers**: No No No Yes Yes Yes Yes Yes
-   **FTP discovery, streaming and downloading**: No No Yes Yes Yes Yes Yes Yes
-   **Store FTP server bookmarks**: No No No Yes Yes Yes Yes Yes
-   **Audio Playback via AirPlay**: Yes Yes Yes Yes Yes Yes Yes Yes
-   **Video Playback via AirPlay**: No Yes Yes Yes Yes Yes Yes Yes
-   **Audio Playback via Connector Cables**: No Yes Yes Yes Yes Yes Yes Yes
-   **Video Playback via Connector Cables**: No Partial Yes Yes Yes Yes Yes Yes
-   **Subtitles playback**: No Western languages only Yes Yes Yes Yes Yes Yes
-   **Subtitles Font Customization**: No No Yes Yes Yes Yes Yes Yes
-   **Closed Caption playback**: No No Yes ^\[2\]^ Yes Yes Yes Yes Yes
-   **Teletext subtitles playback**: No No No Yes Yes Yes Yes Yes
-   **Multi-track audio handling**: No Yes Yes Yes Yes Yes Yes Yes
-   **Video Filtering incl. Screen Brightness**: No Yes Yes Yes Yes Yes Yes Yes
-   **Video Cropping and Aspect Ratio variation**: No Yes Yes Yes Yes Yes Yes Yes
-   **Deinterlacing**: No No No Yes Yes Yes Yes Yes
-   **Playback Speed control**: No Yes Yes Yes Yes Yes Yes Yes
-   **Audio/Subtitles delay control**: No No No No No Yes Yes Yes
-   **Repeated playback**: No No No Yes Yes Yes Yes Yes
-   **Gestures based playback control**: No No No Yes Yes Yes Yes Yes
-   **Playback of Audio-only media (mp3, m4a, flac, …)**: No No Yes Yes Yes Yes Yes Yes
-   **Audio Playback in Background**: No Yes Yes Yes Yes Yes Yes Yes
-   **Mini playback view**: No No No No No No No Yes
-   **Playback timer**: No No No No No No Yes Yes
-   **Chapter & title selection**: No No No No No No Yes Yes
-   **10-band equalizer**: No No No No No No Yes Yes
-   **Playback UI Lock**: No No No No No No Yes Yes
-   **Smart Media Library sorting for audio albums and TV shows**: No No Yes Yes Yes Yes Yes Yes
-   **Media Library Search**: No No No No No Yes Yes Yes
-   **Passcode Lock**: No Yes Yes Yes Yes Yes Yes Yes
-   **VoiceOver support**: No Partial Yes Yes Yes Yes Yes Yes
-   **Organize media in folders**: No No No No Yes Yes Yes Yes
-   **Use folders as playlists**: No No No No Yes Yes Yes Yes
-   **Loop playlists**: No No No No No No No Yes
-   **Playback control through headphones, multi-tasking bar or lock screen**: No No No Partial Yes Yes Yes Yes
-   **Playback is paused when headphones are unplugged**: No No No No Yes Yes Yes Yes
-   **WiFi upload and HTTP downloads in background**: No No No No Yes Yes Yes Yes
-   **Support for password protected HTTP streams**: No No No No Yes Yes Yes Yes
-   **Sharing files with further apps**: No No No No No Yes Yes Yes
-   **Custom vlc:// protocol**: No No No No Partial Yes Yes Yes
-   **Support for x-callback-url**: No No No No No Yes Yes Yes
-   **Apple Watch extension**: No No No No No No No Yes
-   **Supported User Interface Languages**: English English, Danish^\[1\]^, Dutch^\[1\]^, Finnish, French, German, Hebrew^\[1\]^, Indonesian, Italian, Japanese, Russian, Simplified Chinese^\[1\]^, Slovak^\[1\]^, Spanish, Turkish^\[1\]^, Ukrainian^\[1\]^, Vietnamese^\[1\]^ Same as 2.0 plus Bosnian, Catalan, Galician, Greek, Hungarian^\[2\]^, Marathi, Portuguese, Slovenian, Swedish^\[2\]^ Same as 2.1.2 plus Czech, Malay, Persian, Spanish (Mexico), Sinhala ^(added\ in\ 2.2.1)^ Same as 2.2.1 plus British English, Latvian, Romanian Same as 2.3 plus Traditional Chinese Same as 2.4 plus Portuguese (Portugal), Portuguese (Brazil), Khmer, Faroese, Belarusian, Serbian (Latin), Tamil, Afrikaans Same as 2.5
-   **iOS 5.1 support**: Yes Yes Yes Yes No No No No
-   **iOS 6.0 support**: No Yes Yes Yes No No No No
-   **iOS 6.1 support**: No Yes Yes Yes Yes Yes Yes Yes
-   **iOS 7.x support**: No Partial Partial Yes Yes Yes Yes Yes
-   **iOS 8.x support**: No No No No Partial Yes Yes Yes
-   **iOS 9.x support**: No No No No No No No Partial

^\[1\]\ Added\ in\ version\ 2.0.2^ ^\[2\]\ Added\ in\ version\ 2.1.2^

#### Media Synchronization

There are multiple ways to synchronize media files to VLC for iOS. Those may be extended in future releases. Streaming without saving files using the limited space available on iOS devices is also supported. See below.

##### iTunes File Sharing

Using iTunes, you can add and delete files to VLC for iOS. Apple provides [excellent documentation for this](http://support.apple.com/kb/HT4094).

##### WiFi Sharing

If your iOS device and your Mac or PC is on the same local WiFi network, you can use WiFi Upload to add files to VLC for iOS' library.

Within VLC for iOS, click the cone button:

This will expose the sidebar menu. Locate the WiFi Sharing menu item. Notice the empty circle indicating the server's off-state and the description "Inactive Server." (Note: in earlier versions of VLC for iOS, you'll see a toggle button.)

Click the item or switch the toggle. A network address will appear in the item:

Enter the network address to your web browser on your Mac or PC on the same local network:

VLC for iOS' WiFi Sharing page will appear. You can drag any file stored on your Mac or PC to this window for immediate upload to your iOS device. Additionally, you can press the "Upload files" button in the top-right to expose a file selector panel in case your files are not reachable by drag and drop. VLC for iOS' WiFi Sharing supports multiple uploads at the same time and will indicate through a progress bar when upload is complete. Because it's VLC after all, you can start the playback of most files on your iOS device as soon as they appear so you don't need to wait until the upload is done. Version 2.4 adds the ability to also download files stored on device through this page.

##### Dropbox

VLC for iOS natively supports Dropbox. Open the menu as described above, choose Dropbox. For the first time, a login button will appear. When clicking, you'll be redirected to the Dropbox app for login or to the web if you don't have Dropbox installed. VLC for iOS will receive read and write access to your entire Dropbox after login. However, VLC for iOS does not support any write actions (i.e. you cannot upload files from VLC for iOS to your Dropbox either), so don't worry about your file integrity. Starting in version 2.2, VLC for iOS can also stream contents from your Dropbox without downloading them first.

##### Google Drive

Similar to Dropbox, VLC for iOS natively supports Google Drive starting in version 2.2. Please check the process above for setup. Version 2.4 adds support for streaming files from Google Drive without having to download them first.

##### Box.com & OneDrive

Like with Dropbox and Google Drive, version 2.5.0 of VLC for iOS adds support for downloads and direct streaming for both providers.

##### iCloud Drive

With version 2.5.0 of VLC for iOS, any cloud services enabled app including iCloud Drive can be accessed from with the app on iOS 8 and later. While we don't support direct streaming, you can download files from wherever you want without relying on VLC's native implementation.

##### Downloads from the Web

The sidebar menu also includes an item called **Downloads** (or in earlier versions *Download from Web Server*). When selected, it will show the download queue and progress of downloads triggered through the network integration (see below) and exposes a field to enter a URL to directly download media from somewhere. Both HTTP and FTP are supported (earlier versions support *HTTP only*)

#### Network Integration

In addition to the media synchronization options described above, VLC for iOS provides a variety of options to interact with networking sources. Additionally, third-party websites and apps may include links to open streams directly in VLC for iOS.

##### Open Network Stream

Clicking on this item in the sidebar menu reveals a URL field to open a network stream. HTTP, FTP, MMS, MMSH, RTSP, UDP, and RTP streams are supported. Additionally, this view includes a list of your last 15 streams and an option to disable keeping history of your recent streams.

##### Local Network

Introduced in VLC for iOS 2.1, *Local Network* discovers and lists servers found on your local network. At present, this includes UPnP media servers and FTP servers announced via Bonjour / Rendezvous. Further options will be made available in future releases. Depending on the server capability, you can both stream and/or download the stored contents.

Clicking on "Connect To Server" exposes the ability to connect to FTP servers not included in the list.

#### Playback

The controller panel provides access to basic playback controls, a video filter panel, audio and subtitles track selection as well as playback speed. The time slider at the top of the playback screen matches the default media player behavior by allowing you to seek at the pace you want. Next to it, you will find a 2-mode time counter and a button to control aspect ratio and cropping. VLC for iOS will remember the last playback position for media stored on your iOS device.

##### Gestures

Version 2.2 of VLC for iOS introduces gesture-based playback controls in addition to the ordinary buttons.

Tap the playback screen anywhere with 2 fingers to pause or play the current media. Pinch to end the current playback session and close the video.

Swipe to the left and right to change the playback position.

Adapt screen brightness by swiping vertically on the left-hand side of the playback view (gray hands). Change the current volume by swiping vertically on the right-hand side of the playback view (red hands).

##### AirPlay

VLC for iOS supports AirPlay video and audio streaming. To enable audio streaming, just activate the AirPlay switch which will automatically appear next to the volume slider as soon as your iOS device discovers an AirPlay capable playback device (an Apple TV, a multi-media receiver, etc.). For video playback via AirPlay, it's slightly more difficult due to AirPlay API limitations. Apple does not allow to show an AirPlay button for video playback within an Apple if the app does not use the default media player, which VLC does not for the sake of supporting formats other than H264 / MPEG4. As a work-around, you need to use the AirPlay mirroring featuring available from the multi-tasking bar (shown when double-clicking the physical home button on your iOS device).

##### Subtitles and multiple audio tracks

If your media includes subtitles or multiple audio tracks, buttons will appear in the playback controller to switch streams. VLC for iOS will remember the last chosen audio or subtitles track for future playback. If your media does not include subtitles, but you'd like to show some, give them a similar name to the media item, synchronize them and VLC for iOS will discover them automatically.

##### Chapters and titles

With version 2.5.0 and later, you can navigate through your media based on the chapters and titles information includes with Matroska/MKV and MP4 files.

##### Video Filters

Like VLC media player on desktops, VLC for iOS allows you to modify the video's colors in real time. Brightness will adapt your device's physical luminance unless you play your media on an external screen, where it will fallback on a software mode.

##### Playback Speed

Clicking the clock button to the left of the playback controller reveals a slider with a logarithmic scale to adapt the playback speed to your needs. In recent versions, synchronization options for subtitles and audio as well as a playback timer are also available from this menu.

##### Equalizer

In version 2.5.0 or later, a 10-band equalizer is available through the "more" button on the right side of the playback controls. Note that the button will not appear in portrait mode on iPhone.

##### A word on audio playback

VLC for iOS 2.0.x refuses any audio-only media playback. Basic support was added in VLC for iOS 2.1 along with Music Album handling. Future updates will further improve the current feature set by introducing media artwork display, playlists, playlist operations such as shuffling or repeat, and more.

#### Media Library

Your media collection. It offers basic information about each file, such as length, resolution, or file size. Your last playback position is visualized through an orange triangle at the bottom of the snapshot, unless it's new or fully played.

##### Smart handling of music albums and TV shows

VLC for iOS 2.1 added smart handling of music albums and TV shows. Based upon meta tags and pattern matched file names, VLC for iOS will automatically detect TV shows and music albums. Switching the library mode in the sidebar menu will reveal dynamic collections for either category. "All Files" switches back to the default mode listing all files available on your iOS device within the VLC context.

How are TV shows detected by VLC for iOS?—at present, based upon the file name. The following schemes are supported in current releases: "Show.Name.S01E01.Optional.Episode.Name" or "Show.Name.0x00.Optional.Episode.Name". Show Name will become optional in version 2.2.1. Detection for more schemes will be part of future releases.

##### Media grouping in folders

VLC for iOS 2.3 adds support for folders and custom grouping of your media content. A folder also acts as a playlist. In the latest version (2.4.1) you can drag and drop files within a folder, but folders cannot be dragged and dropped. The files or folders cannot be listed automatically in alphabetical order. However, these functions may change as the latest app becomes more stable.

##### Passcode lock

You can lock the app with a passcode. The current versions of VLC for iOS will always ask for it whenever the app is pushed to the foreground. Additionally, your library contents will be hidden away from the multi-tasking switcher. Starting with version 2.3, playback is stopped if passcode lock is enabled and the app is moved to the background to protect your privacy.

###### I forgot my passcode

If you forget your passcode, you need to delete the application and re-install it. This will reset both the settings and the media library. There is no way to recover it. However, you can backup your files using iTunes first. To back up your files using iTunes, sync your iOS device with your computer.

###### I want to use another passcode

Disable passcode lock in VLC's Settings and re-enable it. It will ask you to enter a new passcode.

#### Customization and Settings

VLC for iOS offers a growing variety of options to customize the app suiting your purposes.

-   **Option name**: Version 2.0.x Version 2.1.x Version 2.2 Version 2.3 Version 2.4 Version 2.5 Version 2.6 Details
-   **Passcode Lock**: Yes Yes Yes Yes Yes Yes Yes When enabled, VLC for iOS will ask for the passcode 5 min after leaving the app when using version 2.0 or 2.1. In 2.2, the app will ask for it right away.
-   **Optimize item names for display**: No No Yes Yes Yes Yes Yes Disable removal of unneeded characters, when shown within the media library.
-   **Network caching level**: No No Yes Yes Yes Yes Yes Adapt the network caching level to your needs.
-   **Control playback with gestures**: No No No Yes Yes Yes Yes Disable playback gestures if desired.
-   **Default playback speed**: No No No No Yes Yes Yes
-   **Play video in fullscreen**: No No No No No No Yes ^\[2\]^ On by default. When disabled, video plays minimized by default.
-   **Deblocking filter**: Yes Yes Yes Yes Yes Yes Yes Switch deblocking aggression level. Trade-off between quality and speed.
-   **Deinterlace**: No Yes Yes Yes Yes Yes Yes Deinterlace video image: always or never.
-   **Subtitles Font**: No Yes Yes Yes Yes Yes Yes
-   **Relative Subtitles Font size**: No Yes Yes Yes Yes Yes Yes
-   **Use Bold Font**: No No No Yes Yes Yes Yes
-   **Subtitles Font Color**: No Yes Yes Yes Yes Yes Yes
-   **Text Encoding**: Yes Yes Yes Yes Yes Yes Yes Set the subtitles text encoding mostly used by you
-   **Audio time-stretching**: Yes Yes Yes Yes Yes Yes Yes Improves listening experience
-   **Audio playback in background**: Yes Yes Yes Yes Yes Yes Yes Audio playback continues when leaving VLC for iOS
-   **Unlink from Dropbox**: Yes Yes Yes Yes Yes No ^\[1\]^ No Unlink the app from your Dropbox account
-   **Unlink from Google Drive**: No No Yes Yes Yes No ^\[1\]^ No Unlink the app from your Google Drive account
-   **IPv6 support for WiFi sharing**: No No No No Yes Yes Yes Off by default.
-   **Text Encoding for FTP Connections**: No No No No Yes Yes Yes

^\[1\]\ Version\ 2.5.0\ moves\ those\ buttons\ to\ the\ Cloud\ Services\ UI^ ^\[2\]\ Added\ in\ version\ 2.6.2^

#### Integration with third party apps (version 2.4 or later)

##### Share button

Click the share button within the media library to open stored media in compatible apps on your device. This can be different players, cloud storage or email clients. Availability depends on the support of the other apps.

##### x-callback-url

This is a defined protocol to play or download media in VLC and optionally to go back to the requesting app once playback is done.

    vlc-x-callback://x-callback-url/ACTION?url=...&PARAMETER=...

*Actions:*

**stream**: VLC plays the stream provided by the URL parameter

**download**: VLC will download the file provided by the URL parameter

*Optional Parameters:*

**filename**: VLC will store the file under the given filename when using the **download** action.

**x-success**: VLC will open another x-callback-url once playback is done.

**x-error**: VLC will open another x-callback-url if playback fails. Requires version 2.5.0 or later

#### Installation on iOS 5.1

Starting with version 2.2 of VLC for iOS, we no longer support iOS 5.1. VLC requires iOS 6.1 or later.

### Ubuntu Phone {#ubuntu-phone}


## Miscellaneous {#miscellaneous}

### History {#history}

#### Overview of the VideoLAN project

VideoLAN was a complete software solution for video streaming and playback, developed by students of the [Ecole Centrale Paris](http://www.ecp.fr) and developers from all over the world, under the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) (GPL).

Originally VideoLAN was designed to stream MPEG videos on high-bandwidth networks, but VideoLAN's main software, VLC media player, has evolved to become a full-featured, cross-platform media player.

Now the Non-Profit Organisation developing and offering the VLC media player is called: VideoLAN Organisation

More details about the project can be found on the [VideoLAN Web site](https://www.videolan.org/).

#### VLC Media Player

VLC 2.0 default interface, Windows

Originally called *VideoLAN Client*, VLC media player is VideoLAN's main software product.

VLC media player works on many platforms: Linux, Windows, macOS, BeOS, BSD, Solaris, Android, iOS, QNX and many more... It supports the following video and audio formats: MPEG-1, MPEG-2, MPEG-4/DivX, h264, webm, mkv, DVDs, VCDs, Audio CDs, wmv and wma.

It can also play from external sources:

-   Satellite.
-   Cable.
-   Digital TV cards (DVB-S, DVB-T).
-   Several types of network streams: UDP/RTP Unicast, UDP/RTP Multicast, HTTP, RTSP, MMS, etc.
-   Acquisition or encoding cards.
-   Webcams and other devices.

VLC can also be used as a streaming server. This feature is described in the [Streaming HowTo](#streaming-howto).

This guide describes all the playback (client) aspects of VLC media player.

**Permission is granted to copy, distribute and/or modify this document under the terms of the [GNU General Public License](https://www.gnu.org/copyleft/gpl.html) as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.**
