# Repository guidance

## Commit messages

Use the conventional Git message layout: a concise, type-prefixed summary
line, followed by one blank line, followed by a detailed body when needed.
The summary is the overview of the change; the body explains the reason and
the material implementation details. Select a type that makes the change
category clear, such as `Fix` for bug fixes, `Add` for new functionality,
`Improve` for behavioral improvements, `Docs`, `Test`, `Chore`, or `Verify`.
Make the summary name the affected behavior and outcome; avoid vague messages
that require reading the diff to understand it.

Keep the summary to 50 characters or fewer when practical (never more than 72).
When a body is needed, enter it as real separate lines and separate it from
the summary with exactly one blank line. Wrap body paragraphs at about 72
characters. Do not represent paragraph breaks with escape notation. Omit the
body when the summary alone is sufficient.
