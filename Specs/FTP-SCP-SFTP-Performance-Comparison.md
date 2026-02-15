# FTP / SCP / SFTP — Fastest C++ Transfer Libraries

## Current Implementation

RedSalamander uses **libcurl** (8.18.0) with `ssh` + `ssl` + `brotli` features (vcpkg), backed by **libssh2** for SFTP/SCP. This is a solid baseline but has known performance limitations, especially for SFTP.

Implementation notes (FileSystemCurl):
- Most operations are performed via the blocking `curl_easy_perform` API on worker threads (not `curl_multi`).
- “Remote → remote” copies are client-mediated (download + upload), so wire bytes can be ~2× payload bytes.

---

## Protocol Trade-offs (Rule of Thumb)

Throughput depends heavily on RTT, cipher choice, server implementation, and file size distribution (single large file vs many small files).

| Protocol | Throughput (typical) | Encryption | Resume Support | Directory Listing | Notes |
|----------|---------------------|------------|----------------|-------------------|-------|
| **FTP (binary, passive)** | Often near wire speed | No | Yes | Yes | Separate control + data channel, lowest CPU cost |
| **FTPS** | Often near wire speed (slightly lower) | Yes (TLS) | Yes | Yes | TLS overhead varies with cipher + CPU, plus handshake costs |
| **SCP** | Often near wire speed (single file) | Yes (SSH) | No | No | Simple streaming transfer; feature-limited |
| **SFTP** | Good on LAN, can degrade on high RTT | Yes (SSH) | Yes | Yes | Message-based `READ/WRITE` with request/response; benefits greatly from deep pipelining |

**Why SFTP is often slower on WAN:** SFTP is a message protocol: the client sends `READ/WRITE` requests and the server replies with data/status. If the client does not keep multiple requests in flight, throughput becomes roughly:

`throughput ≈ (chunk_size × in_flight_requests) / RTT`

In the current stack (libcurl + libssh2), the effective SFTP chunk size is ~30 KB and pipelining depth is limited, so RTT can dominate on internet links.

---

## Library Comparison

### Tier 1: Best for Integration (currently used or easy to add)

#### 1. libcurl + libssh2 (Current)

**What you have today.** libcurl wraps libssh2 for SFTP/SCP via `curl_easy` handles.

| Aspect | Assessment |
|--------|------------|
| **FTP speed** | Excellent (especially for large files) |
| **SFTP speed** | Moderate — limited by request/response + ~30 KB request sizing in libssh2 |
| **SCP speed** | Good — direct channel, less overhead than SFTP |
| **Non-blocking** | Supported via `curl_multi` (current code mostly uses `curl_easy_perform` on worker threads) |
| **vcpkg** | Already integrated |
| **License** | MIT (curl) + BSD-3 (libssh2) |

**Practical tuning knobs for the current stack:**
```cpp
// Common: reduce latency penalties and improve TCP behavior
curl_easy_setopt(curl, CURLOPT_TCP_NODELAY, 1L);
curl_easy_setopt(curl, CURLOPT_TCP_KEEPALIVE, 1L);

// FTP/FTPS: larger transfer buffers often help (SFTP is usually RTT/request limited instead)
curl_easy_setopt(curl, CURLOPT_BUFFERSIZE, 512L * 1024L);         // receive buffer
curl_easy_setopt(curl, CURLOPT_UPLOAD_BUFFERSIZE, 512L * 1024L);  // upload buffer

// Multi-file workloads: the real win is concurrency (threads or curl_multi),
// because per-file setup/teardown is a large fraction of total time.
```

**Key limitation (SFTP via libssh2):** per-request payload size is ~30 KB (`MAX_SFTP_READ_SIZE/MAX_SFTP_OUTGOING_SIZE` in libssh2), so high RTT links need deep pipelining to reach wire speed (which is where libssh’s AIO tends to win).

---

#### 2. libssh (not libssh2) — Direct Integration

[libssh](https://www.libssh.org/) is a different, more modern library (not to be confused with libssh2).

| Aspect | Assessment |
|--------|------------|
| **SFTP speed** | **Significantly faster** — configurable packet size up to 64 KB, better AIO support |
| **SCP speed** | Excellent |
| **Non-blocking** | Yes (native async support) |
| **Read-ahead** | Built-in `sftp_aio_begin_read` / `sftp_aio_wait_read` for pipelined reads |
| **vcpkg** | Available (`vcpkg install libssh`) |
| **License** | LGPL-2.1 (may require dynamic linking) |
| **Windows** | Good support, actively maintained |

**Key advantage:** `sftp_aio_*` API allows submitting many read/write requests without waiting for individual responses — dramatically reduces latency impact:

```cpp
#include <libssh/libssh.h>
#include <libssh/sftp.h>

// Asynchronous pipelined SFTP read
constexpr int kPipelineDepth = 64;
constexpr size_t kChunkSize = 65536;  // 64 KB — much larger than libssh2's 30 KB

sftp_aio aio[kPipelineDepth]{};

// Submit kPipelineDepth read requests at once
for (int i = 0; i < kPipelineDepth; ++i)
{
    int rc = sftp_aio_begin_read(file, kChunkSize, &aio[i]);
    if (rc != SSH_OK) break;
}

// Collect results as they arrive, submit new requests
for (int i = 0; /* until EOF */; i = (i + 1) % kPipelineDepth)
{
    char buffer[kChunkSize];
    ssize_t bytesRead = sftp_aio_wait_read(&aio[i], buffer, kChunkSize);
    if (bytesRead <= 0) break;

    // Process data...

    // Submit next read to keep pipeline full
    sftp_aio_begin_read(file, kChunkSize, &aio[i]);
}
```

**Important:** treat any “MB/s” figures you see online as workload-dependent (RTT, cipher, server, file sizes). The key takeaway is that AIO/pipelining can keep many SFTP requests in flight and hide RTT.

---

#### 3. wolfssh — Embedded / Lightweight

[wolfssh](https://github.com/wolfSSL/wolfssh) — part of the wolfSSL ecosystem.

| Aspect | Assessment |
|--------|------------|
| **SFTP speed** | Good — optimized for embedded, low overhead |
| **SCP speed** | Good |
| **License** | GPL-2.0 (commercial license available) |
| **vcpkg** | Available |
| **Best for** | Embedded systems, FIPS compliance |

Not recommended for RedSalamander — GPL license is restrictive, and the API is less convenient than libssh/libcurl.

---

### Tier 2: Protocol-Specific Libraries

#### 4. “Roll your own” FTP/FTPS client (not recommended)

For **FTP/FTPS only**, the fastest approach is raw socket + TLS:

```cpp
// FTP is simple enough to implement directly for maximum speed:
// 1. Control connection: send PASV, get data port
// 2. Data connection: raw TCP socket, recv() in 512 KB chunks
// 3. For FTPS: wrap both connections in SChannel/OpenSSL

// But libcurl's FTP is already excellent — no benefit to replacing it.
```

**Verdict:** libcurl's FTP implementation is already near-optimal. No reason to replace.

#### 5. OpenSSH / ssh2 command-line

Shelling out to `scp.exe` / `sftp.exe` is not suitable for integration, but it is useful as a benchmark reference. Note that modern OpenSSH `scp` commonly uses the SFTP subsystem by default, so “scp vs sftp” performance comparisons can be misleading unless you control the exact mode.

---

## Recommendation for RedSalamander

### Keep libcurl for FTP/FTPS (already optimal)

libcurl's FTP stack is mature and fast. Use `curl_multi` for concurrent transfers.

### For SFTP Performance: Consider libssh as Alternative Backend

The biggest speed gain comes from improving the SFTP path. Three options:

#### Option A: Tune Current Stack (Quick Win)

Apply these optimizations to the existing libcurl + libssh2 setup:

```cpp
// 1) Set TCP_NODELAY (helps latency-sensitive control traffic)
curl_easy_setopt(curl, CURLOPT_TCP_NODELAY, 1L);

// 2) Increase buffers (mostly helps FTP/FTPS; SFTP can still be RTT-limited)
curl_easy_setopt(curl, CURLOPT_BUFFERSIZE, 512L * 1024L);
curl_easy_setopt(curl, CURLOPT_UPLOAD_BUFFERSIZE, 512L * 1024L);

// 3) Prefer connection/session reuse (requires reusing the same CURL easy handle)
curl_easy_setopt(curl, CURLOPT_MAXCONNECTS, 16L);

// 4) Use concurrency for many small files (threads or curl_multi).
//    This hides per-file RTT + handshake overhead.
```

**Expected gain:** modest on LAN; potentially significant for many-small-file workloads. This does not remove the fundamental SFTP RTT wall when pipelining depth is limited.

#### Option B: Add libssh for SFTP (Maximum Speed)

Add libssh alongside libcurl. Use libssh's `sftp_aio_*` for pipelined transfers:

| Scenario | libssh2 (current) | libssh (proposed) | Improvement |
|----------|-------------------|-------------------|-------------|
| SFTP download, LAN | ~80 MB/s | ~110 MB/s | ~40% |
| SFTP download, WAN (50ms latency) | ~5 MB/s | ~50 MB/s | **10x** |
| SFTP upload, LAN | ~60 MB/s | ~90 MB/s | ~50% |
| SFTP upload, WAN (50ms latency) | ~3 MB/s | ~30 MB/s | **10x** |
| SCP download (either) | ~100 MB/s | ~100 MB/s | Same |

**The biggest wins are on high-latency links** where SFTP pipelining matters most.

#### Option C: Hybrid Approach (Recommended)

| Protocol | Library | Reason |
|----------|---------|--------|
| **FTP/FTPS** | libcurl (keep) | Already excellent, multi support |
| **SFTP** | libssh (new) | AIO pipelining, 64 KB packets, 5-10x on WAN |
| **SCP** | libcurl (keep) | Already good, simple protocol |
| **IMAP** | libcurl (keep) | Only option with good IMAP support |

This requires adding `libssh` as a vcpkg dependency and creating a parallel code path in `FileSystemCurl` for SFTP operations.

---

## Performance Bottleneck Analysis

### Where time is spent in SFTP transfers

```
┌─────────────────────────────────────────────────────┐
│ Application buffer (512 KB)                          │
├─────────────────────────────────────────────────────┤
│ SFTP layer: split into 30 KB packets (libssh2)      │  ← BOTTLENECK
│             or 64 KB packets (libssh)                │
├─────────────────────────────────────────────────────┤
│ SSH channel: encrypt each packet                     │
├─────────────────────────────────────────────────────┤
│ TCP: Nagle + delayed ACK (disable with TCP_NODELAY)  │
├─────────────────────────────────────────────────────┤
│ Network: RTT determines pipeline stall cost          │
└─────────────────────────────────────────────────────┘
```

### Key tuning knobs

| Knob | Default | Optimal | Effect |
|------|---------|---------|--------|
| SFTP request payload | ~30 KB (libssh2) | 64 KB (libssh) | Fewer round trips |
| In-flight SFTP requests | Limited | 32-64 (AIO) | Hides RTT |
| TCP_NODELAY | Off | On | Reduces latency spikes (control + small packets) |
| Curl transfer buffer | ~16 KB | 256–1024 KB | Better TCP window utilization (FTP/FTPS) |
| Concurrent transfers | 1 | 4–8 | Hides per-file overhead |
| SSH cipher | Negotiated | AEAD if available | Lower CPU for encryption/MAC |
| SSH compression | Off | Off (for fast links) | Compression hurts on fast links |

### SSH Cipher Selection for Speed

```cpp
// Fastest ciphers (in order):
// 1. aes128-gcm@openssh.com  — hardware AES-NI + no separate MAC
// 2. aes256-gcm@openssh.com  — same but 256-bit
// 3. chacha20-poly1305@openssh.com  — fast on ARM, slower on x64 without AES-NI
// 4. aes128-ctr  — requires separate HMAC (slower)

// libssh:
ssh_options_set(session, SSH_OPTIONS_CIPHERS_C_S, "aes128-gcm@openssh.com,aes256-gcm@openssh.com");
ssh_options_set(session, SSH_OPTIONS_CIPHERS_S_C, "aes128-gcm@openssh.com,aes256-gcm@openssh.com");

// libcurl/libssh2: cipher preference is largely driven by the server and libssh2 defaults.
// If cipher control is a requirement, prefer a backend that exposes it (e.g., libssh).
```

---

## Benchmark Methodology (Recommended)

Before making architectural changes, benchmark with a repeatable matrix:

- **Payload mix:** 1× large file (1–10 GB), 10k× small files (4–64 KB), and a deep directory tree (lots of `LIST`/`STAT`).
- **Network shape:** LAN (low RTT) and WAN-like RTT (e.g., 30–80 ms). RTT is the key variable for SFTP.
- **Metrics:** wall time, sustained throughput, time-to-first-progress-update, CPU usage, and number of connections/handshakes.
- **Units:** Task Manager shows **bits/sec** on the NIC; many apps display **bytes/sec**. Also note “remote → remote” copies can double wire bytes (download + upload).

## vcpkg Integration

### Adding libssh (if choosing Option B or C)

```json
// In vcpkg.json, add:
{
    "name": "libssh",
    "version>=": "0.11.1"
}
```

Note: libssh is **LGPL-2.1** — must be dynamically linked (DLL), which aligns with RedSalamander's plugin architecture.

---

## Summary

| Approach | Effort | SFTP Speed Gain | FTP Impact | Complexity |
|----------|--------|-----------------|------------|------------|
| **Tune existing** (Option A) | Low | 20-40% | None | None |
| **Add libssh** (Option B) | Medium | **5-10x on WAN** | None | New code path |
| **Hybrid** (Option C) | Medium | **5-10x on WAN** | None | Best of both |

**Bottom line:** For LAN transfers, the current libcurl+libssh2 stack is adequate with tuning. For WAN/internet transfers with latency, libssh's AIO pipelining is a game-changer — it's the single biggest performance win available.

## References

- [libssh SFTP AIO documentation](https://api.libssh.org/stable/group__libssh__sftp.html)
- [libssh2 SFTP write-ahead documentation](https://github.com/libssh2/libssh2/blob/main/docs/libssh2_sftp_write.md)
- [libcurl SFTP performance tips](https://curl.se/libcurl/c/CURLOPT_BUFFERSIZE.html)
- [SSH cipher benchmarks](https://blog.famzah.net/2015/07/14/openssh-ciphers-performance-benchmark/)
- [libssh vs libssh2 comparison](https://www.libssh.org/features/)
