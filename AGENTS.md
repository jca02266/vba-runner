# Repository guidance

## Commit messages

Use a type-prefixed, concise one-line summary. Select a type that makes the
change category clear, such as `Fix` for bug fixes, `Add` for new functionality,
`Improve` for behavioral improvements, `Docs`, `Test`, `Chore`, or `Verify`.
Make the summary name the affected behavior and outcome; avoid vague messages
that require reading the diff to understand the change.

Keep the summary to 50 characters or fewer when practical (never more than 72).
When a body is needed, separate it from the summary with one blank line, wrap
paragraphs at about 72 characters, and write actual line breaks—never literal
`\\n` sequences. Use the body for the reason and material implementation details;
omit it when the summary alone is sufficient.
