# I66 Skip before publishing a Move with metadata loss

The fixture moves a downloaded file with Mark of the Web into an occupied name through
the owned Local loopback path. It accepts the replacement question and chooses Skip when
the stage reports loss of the named stream.

The first x64 Debug candidate `4690964` reached both questions in
`20261125T103555Z-17176-8b7ce792b0df40ccb26089843f31c571`. It incorrectly returned
aggregate success for the unpublished item. The repair counts the explicit decision so the
task reduces the result to Skipped.

Integrated `0ac38620b` built x64 Debug with zero warnings and errors. Receipt
`a5d239e491f1e882380ad7f80670b5ff0cd4810cb562878ccc5ea05675439d6d` and run
`20261125T113013Z-81044-5542320ea9104f68889e3984ad2dcc81` passed all three focused stages
without desktop activation. The source and named stream remained byte-exact. The occupied
destination stayed byte-exact and the unpublished stage was removed. The typed outcome
recorded Skipped, NotPublished, Retained and Removed. `FileOps.Metadata.StageSkipped`
recorded the Mark-of-the-Web mask and the transaction recorded no publication.

The same run checked a direct-final Move whose destination was already published. It retained
the source and verified the published copy without another question. The case qualifies these
two local paths. Predictable sensitive loss and other provider routes use separate evidence.
