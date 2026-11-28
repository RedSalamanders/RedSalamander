# Third-party notices

RedSalamander is copyright (c) 2026 RedSalamander, under the MIT license in [LICENSE.txt](LICENSE.txt). This file lists
the third-party components shipped with RedSalamander and identifies their notices. Add a new component and its full
notice before distribution. The vcpkg entries reflect the pinned registry baseline and overrides in `vcpkg.json`.
A `#` suffix identifies a vcpkg port revision. The dependency and terminal runtime locks identify the pinned sources.

## Third-Party Licenses

The entries below identify the components, applicable terms and available notice texts.

-------------------------------------------------------------------------------

### Dracula theme palette

- Snapshot: c988d3d1c9e40dffef7c28f4b99079523637d5d2
- License: MIT
- Copyright (c) 2023 Dracula Theme
- Exact license text: `Themes\Licenses\Dracula.LICENSE.txt`

https://github.com/dracula/dracula-theme

-------------------------------------------------------------------------------

### Catppuccin palette

- Version: 1.8.0 (commit 07d02aa110ef9eb7e7427afca5c73ba9cf7f8ebd)
- License: MIT
- Copyright (c) 2021 Catppuccin
- Exact license text: `Themes\Licenses\Catppuccin.LICENSE.txt`

https://github.com/catppuccin/palette

-------------------------------------------------------------------------------

### WIL (Windows Implementation Libraries)

- Version: 1.0.260126.7
- License: MIT
- Copyright (c) Microsoft Corporation. All rights reserved.

https://github.com/microsoft/wil

-------------------------------------------------------------------------------

### yyjson

- Version: 0.13.0
- License: MIT
- Copyright (c) 2020 YaoYuan <ibireme@gmail.com>

https://github.com/ibireme/yyjson

-------------------------------------------------------------------------------

### pe-parse

- Version: 2.1.1#1
- License: MIT
- Copyright (c) 2013 Andrew Ruef

https://github.com/trailofbits/pe-parse

-------------------------------------------------------------------------------

### libjpeg-turbo

- Version: 3.2.0#1
- License: IJG License AND BSD-3-Clause (modified 3-clause BSD)
- Copyright (C) 2009-2025 D. R. Commander. All Rights Reserved.
- Copyright (C) 2015 Viktor Szathmary. All Rights Reserved.

This software is based in part on the work of the Independent JPEG Group.

https://github.com/libjpeg-turbo/libjpeg-turbo

-------------------------------------------------------------------------------

### LibRaw

- Version: 0.22.2
- License: (LGPL-2.1-only OR CDDL-1.0) AND BSD-3-Clause
- Copyright (C) LibRaw LLC

https://www.libraw.org/

-------------------------------------------------------------------------------

### 7-Zip

- Version: 26.5
- License: GNU LGPL-2.1-or-later (most files), with BSD-3-Clause and
  BSD-2-Clause for some components
- Copyright (C) 1999-2026 Igor Pavlov

Additional copyrights for bundled components:
- LZFSE: Copyright (c) 2015-2016 Apple Inc. (BSD-3-Clause)
- ZSTD: Copyright (c) Facebook, Inc. (BSD-3-Clause); Copyright (c) 2023-2026 Igor Pavlov
- XXH64: Copyright (c) 2012-2021 Yann Collet; Copyright (c) 2023-2026 Igor Pavlov (BSD-2-Clause)

Note: RAR-related files (CPP/7zip/Compress/Rar*) are subject to the
unRAR license restriction (see unRAR below).

https://www.7-zip.org/

-------------------------------------------------------------------------------

### unRAR

- Source: RAR decompression code included with the pinned 7-Zip port
- License: unRAR freeware license
- All copyrights to RAR and the utility UnRAR are exclusively owned by
  Alexander L. Roshal.

UnRAR source code may be used in any software to handle RAR archives without
limitations free of charge, but cannot be used to develop RAR (WinRAR)
compatible archiver and to re-create RAR compression algorithm, which is
proprietary. Distribution of modified UnRAR source code in separate form or as
a part of other software is permitted, provided that full text of this
paragraph, starting from "UnRAR source code" words, is included in license, or
in documentation if license is not available, and in source code comments of
resulting package.

https://www.rarlab.com/

-------------------------------------------------------------------------------

### curl

- Version: 8.23.0
- License: curl, ISC and BSD-3-Clause terms
- Copyright (c) 1996-2025 Daniel Stenberg <daniel@haxx.se>, and many
  contributors.

Permission to use, copy, modify, and distribute this software for any purpose
with or without fee is hereby granted, provided that the above copyright notice
and this permission notice appear in all copies.

https://curl.se/

-------------------------------------------------------------------------------

### FTXUI

- Version: 7.0.3
- License: MIT
- Copyright (c) 2019 Arthur Sonzogni.

https://github.com/ArthurSonzogni/FTXUI

-------------------------------------------------------------------------------

### libssh2

- Version: 1.11.1#3
- License: BSD-3-Clause
- Copyright (C) 2004-2007 Sara Golemon <sarag@libssh2.org>
- Copyright (C) 2005-2006 Mikhail Gusarov
- Copyright (C) 2006-2007 The Written Word, Inc.
- Copyright (C) 2007 Eli Fant
- Copyright (C) 2009-2023 Daniel Stenberg
- Copyright (C) 2008-2009 Simon Josefsson
- Copyright (C) 2000 Markus Friedl
- Copyright (C) 2015 Microsoft Corp.

https://libssh2.org/

-------------------------------------------------------------------------------

### OpenSSL

- Version: 3.6.5#2
- License: Apache-2.0
- Copyright (c) 1998-2025 The OpenSSL Project Authors. All Rights Reserved.

https://www.openssl.org/

-------------------------------------------------------------------------------

### Brotli

- Version: 1.2.0
- License: MIT
- Copyright (c) 2009, 2010, 2013-2016 by the Brotli Authors.

https://github.com/google/brotli

-------------------------------------------------------------------------------

### zlib

- Version: 1.3.2#2
- License: zlib License
- Copyright (C) 1995-2022 Jean-loup Gailly and Mark Adler

https://zlib.net/

-------------------------------------------------------------------------------

### JasPer

- Version: 4.2.9
- License: JasPer-2.0 (MIT-like permissive)
- Copyright (c) 2001-2016 Michael David Adams
- Copyright (c) 1999-2000 Image Power, Inc.
- Copyright (c) 1999-2000 The University of British Columbia

https://www.ece.uvic.ca/~frodo/jasper/

-------------------------------------------------------------------------------

### Little CMS (lcms2)

- Version: 2.19.1#1
- License: MIT
- Copyright (c) 2023 Marti Maria Saguer

https://www.littlecms.com/

-------------------------------------------------------------------------------

### WebView2

- Version: 1.0.4258.31
- License: BSD-3-Clause
- Copyright (C) Microsoft Corporation. All rights reserved.

https://developer.microsoft.com/en-us/microsoft-edge/webview2/

-------------------------------------------------------------------------------

### AWS SDK for C++

- Version: 1.11.880
- License: Apache-2.0, MIT and Zlib terms
- Copyright Amazon.com, Inc. or its affiliates. All Rights Reserved.

The vcpkg manifest overrides the AWS SDK and AWS CRT C++ versions. The associated AWS C Runtime libraries are:

- aws-crt-cpp 0.43.4
- aws-c-common 1.0.3
- aws-c-s3 1.3.1
- aws-c-auth 1.0.1
- aws-c-cal 1.0.0
- aws-c-io 1.1.0
- aws-c-http 1.1.1
- aws-c-mqtt 1.1.2
- aws-c-compression 1.0.0
- aws-c-event-stream 1.0.0
- aws-c-sdkutils 1.0.0
- aws-checksums 1.0.1

https://github.com/aws/aws-sdk-cpp

-------------------------------------------------------------------------------

The portable ZIP carries the full ViewerWeb license and jsoneditor NOTICE texts in `Plugins/Licenses/ViewerWeb/`.

### highlight.js (ViewerWeb plugin)

- Version: 11.11.1
- License: BSD-3-Clause
- Copyright (c) 2006-2024 Josh Goebel and other contributors.

https://highlightjs.org/

-------------------------------------------------------------------------------

### jsoneditor (ViewerWeb plugin)

- Version: 10.4.2
- License: Apache-2.0
- Copyright (c) Jos de Jong

https://github.com/josdejong/jsoneditor

-------------------------------------------------------------------------------

### markdown-it (ViewerWeb plugin)

- Version: 14.1.0
- License: MIT
- Copyright (c) 2014 Vitaly Puzrin, Alex Kocharin.

https://github.com/markdown-it/markdown-it

-------------------------------------------------------------------------------

### BLAKE3

- Version: 1.8.7
- License: CC0-1.0 OR Apache-2.0

[BLAKE3 source and license files](https://github.com/BLAKE3-team/BLAKE3/tree/1.8.7)

-------------------------------------------------------------------------------

### SQLite

- Version: 3.54.1
- Terms: Public-domain dedication for the SQLite deliverable code

[SQLite copyright and public-domain dedication](https://sqlite.org/copyright.html)

-------------------------------------------------------------------------------

### nghttp2

- Version: 1.70.0
- License: MIT

[nghttp2 copyright and license](https://github.com/nghttp2/nghttp2/blob/v1.70.0/COPYING)

-------------------------------------------------------------------------------

### DxUi

- Source: `Dependencies/DxUi.lock.json`
- License: MIT
- Copyright (c) 2026 RedSalamander

DxUi compiles the Unicode 17.0.0 tables under the Unicode License V3 and WHATWG named characters under the
BSD 3-Clause and CC BY 4.0 terms. [Pinned DxUi notices](https://github.com/RedSalamanders/DxUi/blob/7fb2ab71852087d93b3a3e5b3ba504c37f56267c/THIRD-PARTY-NOTICES.md)

-------------------------------------------------------------------------------

### RedMonitor client

- Source: `Dependencies/RedMonitor.lock.json`
- License: MIT
- Copyright (c) 2026 RedSalamander
- Packaged notice: `RedMonitor-LICENSE.txt`

-------------------------------------------------------------------------------

### Ghostty terminal runtime

The runtime source and each bundled component notice are pinned by
`External/TerminalEngine/GhosttyRuntimeLock.v1.json`. The runtime retains the notice texts listed under
`share/licenses/terminal-runtime/` in that lock. [Pinned Ghostty source and license](https://github.com/ghostty-org/ghostty/tree/83fb95624b0439ee660c2679b43b38c8b63769cd)

-------------------------------------------------------------------------------

### Microsoft Visual C++ runtime

The app-local runtime DLLs are supplied under the Microsoft Visual Studio distributable-code terms.
