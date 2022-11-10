- Make sure `expandTrusted`, `convertToString`, `parseDOM` are only called once
  (and cached) for every ArchiveMember. Cache that calculation in an AA mapping
  from `ArchiveMember` to `RelationshipsById`.
- Move body of `parseRelationships` into `File.parseRelationships` and make the
  module-scope public `parseRelationships` a thin wrapper on top of it.
