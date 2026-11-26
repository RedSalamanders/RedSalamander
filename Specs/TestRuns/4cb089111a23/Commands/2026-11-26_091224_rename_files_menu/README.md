# Scheduled Rename in the Files menu

The test-enabled x64 Release build at `4dc5bd30a6fd94ca76981046a18cb3780305b0b6` completed with zero warnings and errors. Its verified receipt is `8dda06909ca68de6d449fc8c89213f35dbbc41fae4d269d4bf9bd54347977b38`.

Run `20261126T081147Z-57352-f7761f13f0eb4f57afe6ec60347254aa` selected only `main_menu_information_architecture`. The guarded Commands runner reported zero pending desktop entries and passed one case with zero failures and skips in 94 ms. It initialized the Files menu by message, exercised the direct command, restored the pane and retired its two owned child processes. The focused runner reported repository `NOT_EVALUATED`. The disk audit listed only existing entries outside this run.

| Provider | Batch Rename and Change Case | Menu refresh |
| --- | --- | ---: |
| Dummy | Disabled | 27 µs |
| Local | Enabled | 16 µs |

Direct Batch Rename on Dummy returned `ERROR_NOT_SUPPORTED` in a visible nonblocking pane alert. The menu refresh metric measured two command states per provider switch. The exact no-activation admission also passed the Product documentation capture fixture group and the 67 selected runner-plan checks for exact filters, proof hooks and desktop classification. The broader source check found the obsolete synchronous Pack assertion, preserved separately in `20261126T081430Z-66104-24d749477c74448da3027f5553ce044c`. Its correction checks worker publication and the retained archive handle before cleanup admission.
