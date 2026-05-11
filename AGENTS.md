# Agent Information

## How to Start the Web App (for Development)

If you need to verify operations on the Web UI or test editor features, start the development server using the following commands:

```bash
# Install dependencies (only required the first time)
npm install

# Start the local development server (default: http://localhost:5173/)
npm run dev
```

## How to Run Automated Tests

This project includes a TypeScript-based unit testing environment for verifying the logic of VBA code.
When an agent makes changes to the code, please run the following command to ensure there are no regressions:

```bash
# Bundle and run tests (AST construction and verification)
npx esbuild sample/tests/ts/TaskScheduler_Core.test.ts --bundle --outfile=sample/tests/ts/TaskScheduler_Core.test.cjs --platform=node && node sample/tests/ts/TaskScheduler_Core.test.cjs
```

## File I/O Sandbox Policy

When implementing or modifying file-related functions (e.g., `Open`, `Kill`, `Dir`, `Environ`), you **MUST** follow the Sandbox rules described in `README.md`:

- **All file paths** must be resolved through the `SandboxPath` class in `src/compiler/sandbox.ts`.
- **Sandbox Root**: File operations are restricted to the sandbox root (default: `workspace/`).
- **No Traversal**: Any attempt to access paths outside the sandbox root (e.g., using `../`) must result in a runtime error.
- **Path Virtualization**: Windows-style absolute paths are mapped to subdirectories under the sandbox root (e.g., `C:\foo` -> `{sandboxRoot}/c/foo`).
- **Environment Variables**: The `Environ` function must only access the sandbox's virtual environment, not the host OS environment variables.

## Environment-Dependent Operations Policy

This project's primary goal is **Refactoring and Unit Testing**. Therefore, all operations that depend on the host environment (external commands, GUI, network, etc.) must be implemented as **mocks or dummies**.

- **External Commands**: Functions like `Shell` must not execute actual commands. They should log the command and return a success task ID (default: `1`).
- **External Files**: All file operations must be restricted to the sandbox root via `SandboxPath`. Do not access files outside the `workspace/` directory.
- **External Objects**: `CreateObject` and `GetObject` must return mock objects for standard libraries (e.g., `Scripting.Dictionary`, `MSXML2.XMLHTTP`) and not attempt to instantiate actual COM objects or make network requests.
- **Registry**: Functions like `GetSetting` and `SaveSetting` must use a virtual registry (e.g., a memory map) rather than the actual Windows Registry.
- **App State**: Functions like `AppActivate` and `SendKeys` must be implemented as no-ops or log the action.
- **UI/Interactions**: `MsgBox` and `InputBox` must not block execution. They should log the message and return a default value.
