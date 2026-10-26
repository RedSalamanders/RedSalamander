# P2.7 queue-progress timing continuation diagnostic

An intermediate complete File Operations run recorded one failure in
`Phase5_SwitchParallelToWaitDuringDiscovery`: the timing-sensitive queue-progress assertion did not
observe its expected window. The retained `failed_*` files are the complete case/trace evidence from
that run.

The same case passed in the immediately following isolated run retained as `isolated_*`, then passed
again in the subsequent complete 117/0/20 P2.7 validation. No code change was made to suppress or
quarantine the assertion, and the final run reports zero flaky, regression, isolation, or
unclassified-failure classifications. This archive records the observed timing sensitivity rather
than treating the earlier failure as product evidence.
