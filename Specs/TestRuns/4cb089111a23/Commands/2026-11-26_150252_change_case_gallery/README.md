# Change Case documentation gallery

The test-enabled x64 Release candidate `2d41ebc552455a91b5a22df48b0d934957f81200` built with zero compiler warnings and errors. Its verified receipt was `a7b64568bc0faeacff88113158590683047c0b3e0b06616e829cee31e136299b`. Both exact Commands selections ran with `--selftest-no-activate`.

| Selection | Archived run | Passed / failed / skipped |
| --- | --- | --- |
| `cmd_product_documentation_gallery`, baseline `78b8c46d3` | `20261126T134952Z-61592-96b7757f7b29461c887b03168259ae27` | 0 / 1 / 0 |
| `cmd_product_documentation_gallery`, corrected | `20261126T140159Z-47256-51aadb8760f54c67a77df0a00d8be6a4` | 1 / 0 / 0 |
| `main_menu_information_architecture`, corrected | `20261126T140402Z-71920-56605701ba664fc784e989ba0782858f` | 1 / 0 / 0 |

The baseline failed because the Dummy demonstration pane could no longer open a scheduled Change Case command after exact Rename admission was enforced. The corrected gallery uses the guarded presentation entry point. It recorded 45 scenarios with `noActivation=true` and captured `change-case.png` at 144 DPI and 694 × 512 pixels. The report is in `commands/product-documentation/scenarios.tsv` alongside the captures. The Files-menu case continued to verify the scheduled Rename command states on Dummy and Local.

The initial failure also appears in the [hosted regression on `20333bca6`](https://github.com/RedSalamanders/RedSalamander/actions/runs/45791770365). The downloaded x64 and ARM64 artifacts remain archived as `20261126T134905Z-60780-d936cf4f136946ad9b097e81ab87173b` and `20261126T134905Z-67296-b3df7fdb34e946d5a8168df118df1f52`. This run addresses the gallery failure. Broader native qualification, including the interactive pending-decision disclosure case, is still open.
