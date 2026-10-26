// X2 model: ViewerSpace scan publication, records posted while holding the frontier mutex (old, 1dcd28c)
// versus after releasing it (new, 3fa79cf). This is a threading model, not the plugin: enumeration is a busy
// spin, the frontier is a job counter, and the UI drain follows DrainUpdates (512 records per 16 ms tick).
// PostUpdate's backpressure thresholds are copied from ViewerSpace::PostUpdate.
#include <algorithm>
#include <atomic>
#include <chrono>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <deque>
#include <mutex>
#include <string>
#include <thread>
#include <vector>

namespace
{
using Clock = std::chrono::steady_clock;

struct Config
{
    uint32_t workers;
    uint32_t directories;
    uint32_t enumerateUs;      // per-directory provider cost
    uint32_t recordsPerDir;    // AddChild + UpdateSize + DirectoryFilesSummary + states, per retained directory
    uint32_t recordBytes;      // EstimatePendingUpdateBytes, including names and file summaries
    bool postUnderLock;        // old behaviour
};

struct Queue
{
    std::mutex mutex;
    std::deque<uint32_t> records;
    std::atomic_uint64_t pendingBytes{0};
};

void PostUpdate(Queue& queue, uint32_t bytes)
{
    uint64_t pending;
    {
        std::scoped_lock lock(queue.mutex);
        queue.records.push_back(bytes);
        pending = queue.pendingBytes.fetch_add(bytes) + bytes;
    }
    constexpr uint64_t kSoft = 16ull * 1024 * 1024;
    constexpr uint64_t kHard = 64ull * 1024 * 1024;
    if (pending > kHard)
        std::this_thread::sleep_for(std::chrono::milliseconds(5));
    else if (pending > kSoft)
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    else if (pending > kSoft / 4)
        std::this_thread::yield();
}

void Spin(uint32_t us)
{
    const auto until = Clock::now() + std::chrono::microseconds(us);
    while (Clock::now() < until)
    {
    }
}

struct Result
{
    double seconds;
    double p50WaitUs, p95WaitUs, maxWaitUs;
    double workerBlockedShare; // share of worker time spent waiting for the frontier lock
};

Result Run(const Config& config)
{
    Queue queue;
    std::mutex frontierMutex;
    uint32_t nextJob = 0;
    std::atomic_bool producersDone{false};
    std::atomic_uint64_t totalWaitNs{0};
    std::vector<std::vector<double>> waits(config.workers);

    std::thread consumer([&]() {
        while (! producersDone.load() || queue.pendingBytes.load() > 0)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(16));
            std::scoped_lock lock(queue.mutex);
            for (int i = 0; i < 512 && ! queue.records.empty(); ++i)
            {
                queue.pendingBytes.fetch_sub(queue.records.front());
                queue.records.pop_front();
            }
        }
    });

    const auto started = Clock::now();
    std::vector<std::thread> workers;
    for (uint32_t w = 0; w < config.workers; ++w)
    {
        workers.emplace_back([&, w]() {
            for (;;)
            {
                {
                    std::scoped_lock lock(frontierMutex); // job pop
                    if (nextJob >= config.directories)
                        return;
                    ++nextJob;
                }
                Spin(config.enumerateUs);

                const auto requested = Clock::now();
                std::unique_lock lock(frontierMutex); // publication
                const auto acquired = Clock::now();
                const double waitUs = std::chrono::duration<double, std::micro>(acquired - requested).count();
                waits[w].push_back(waitUs);
                totalWaitNs.fetch_add(static_cast<uint64_t>(waitUs * 1000.0));
                if (config.postUnderLock)
                {
                    for (uint32_t r = 0; r < config.recordsPerDir; ++r)
                        PostUpdate(queue, config.recordBytes);
                    lock.unlock();
                }
                else
                {
                    lock.unlock();
                    for (uint32_t r = 0; r < config.recordsPerDir; ++r)
                        PostUpdate(queue, config.recordBytes);
                }
            }
        });
    }
    for (auto& t : workers)
        t.join();
    const double seconds = std::chrono::duration<double>(Clock::now() - started).count();
    producersDone.store(true);
    consumer.join();

    std::vector<double> all;
    for (auto& v : waits)
        all.insert(all.end(), v.begin(), v.end());
    std::sort(all.begin(), all.end());
    auto pct = [&](double q) { return all.empty() ? 0.0 : all[std::min(all.size() - 1, static_cast<size_t>(q * all.size()))]; };
    const double blocked = (totalWaitNs.load() / 1e9) / (seconds * config.workers);
    return {seconds, pct(0.50), pct(0.95), all.empty() ? 0.0 : all.back(), blocked};
}
} // namespace

int main(int argc, char** argv)
{
    const uint32_t workers = argc > 1 ? static_cast<uint32_t>(std::atoi(argv[1])) : 4u;
    struct Scenario
    {
        const char* name;
        uint32_t directories, enumerateUs, recordsPerDir, recordBytes;
    };
    // "light": records small enough that the queue never reaches the 4 MiB yield threshold.
    // "backpressure": warm metadata and ~1 KiB records (names + file summaries) push the queue past 16 MiB.
    const Scenario scenarios[] = {{"light", 20'000, 50, 6, 96}, {"backpressure", 20'000, 30, 8, 1024}};
    std::printf("scenario,workers,variant,dirs_per_s,lock_wait_p50_us,lock_wait_p95_us,lock_wait_max_us,worker_time_blocked_on_lock\n");
    for (const Scenario& s : scenarios)
    {
        for (bool underLock : {true, false})
        {
            const Config config{workers, s.directories, s.enumerateUs, s.recordsPerDir, s.recordBytes, underLock};
            const Result r = Run(config);
            std::printf("%s,%u,%s,%.0f,%.1f,%.1f,%.0f,%.1f%%\n",
                        s.name,
                        workers,
                        underLock ? "old_post_under_lock" : "new_post_after_unlock",
                        s.directories / r.seconds,
                        r.p50WaitUs,
                        r.p95WaitUs,
                        r.maxWaitUs,
                        r.workerBlockedShare * 100.0);
        }
    }
    return 0;
}
