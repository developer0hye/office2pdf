# GitHub CLI pagination

- Implementation: <https://github.com/cli/cli/tree/trunk/pkg/cmd/api>
- API manual: <https://cli.github.com/manual/gh_api>
- `gh api graphql --paginate` follows the connection's `pageInfo` until
  `hasNextPage` is false. The query must accept `$endCursor: String` and select
  both `hasNextPage` and `endCursor`.
- `-F owner='{owner}' -F name='{repo}'` resolves the current repository; other
  field arguments become GraphQL variables.
- Used by `scripts/issue_loop.sh` to enumerate merged/open repository PRs. A
  bounded search can miss older PRs or recently merged PRs absent from the index.
- Capture the complete command output before reporting progress, and propagate
  its failure status so a failed later page cannot become a partial count.
