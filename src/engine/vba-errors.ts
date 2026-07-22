/**
 * VBA runtime error number constants and default messages.
 *
 * Source: MS-VBAL §6.1.2.1 and standard VBA runtime error list.
 * Error 1 (Option Explicit violation) is an engine extension, not in MS-VBAL.
 */

export const VbaErrorCode = {
    // --- Custom engine errors ---
    OPTION_EXPLICIT_VIOLATION:           1,
    CONSTANT_EXPRESSION_REQUIRED:        2,

    // --- Control flow ---
    RETURN_WITHOUT_GOSUB:                3,
    RESUME_WITHOUT_ERROR:               20,
    SUB_OR_FUNCTION_NOT_DEFINED:        35,

    // --- Type / arithmetic ---
    INVALID_PROCEDURE_CALL:              5,
    OVERFLOW:                            6,
    TYPE_MISMATCH:                      13,
    OUT_OF_STRING_SPACE:                14,
    DIVISION_BY_ZERO:                   11,

    // --- Collection / array ---
    SUBSCRIPT_OUT_OF_RANGE:              9,
    KEY_ALREADY_EXISTS:                457,

    // --- File I/O ---
    BAD_FILE_NAME_OR_NUMBER:            52,
    FILE_NOT_FOUND:                     53,
    BAD_FILE_MODE:                      54,
    FILE_ALREADY_OPEN:                  55,
    FILE_ALREADY_EXISTS:                58,
    TOO_MANY_FILES:                     67,
    PERMISSION_DENIED:                  70,
    PATH_FILE_ACCESS_ERROR:             75,
    PATH_NOT_FOUND:                     76,

    // --- Object ---
    OBJECT_VARIABLE_NOT_SET:            91,
    INVALID_USE_OF_NULL:                94,
    OBJECT_REQUIRED:                   424,
    ACTIVEX_CANT_CREATE_OBJECT:        429,
    OBJECT_DOESNT_SUPPORT_PROPERTY:    438,
    ARGUMENT_NOT_OPTIONAL:             449,
    WRONG_NUMBER_OF_ARGUMENTS:         450,
} as const;

/** Default error messages for static (non-dynamic) errors. */
export const VBA_ERROR_MESSAGES: Readonly<Record<number, string>> = {
    [VbaErrorCode.OPTION_EXPLICIT_VIOLATION]:        'Variable not declared (Option Explicit)',
    [VbaErrorCode.RETURN_WITHOUT_GOSUB]:             'Return without GoSub',
    [VbaErrorCode.RESUME_WITHOUT_ERROR]:             'Resume without error',
    [VbaErrorCode.INVALID_PROCEDURE_CALL]:           'Invalid procedure call or argument',
    [VbaErrorCode.OVERFLOW]:                         'Overflow',
    [VbaErrorCode.TYPE_MISMATCH]:                    'Type mismatch',
    [VbaErrorCode.OUT_OF_STRING_SPACE]:              'Out of string space',
    [VbaErrorCode.DIVISION_BY_ZERO]:                 'Division by zero',
    [VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE]:           'Subscript out of range',
    [VbaErrorCode.KEY_ALREADY_EXISTS]:               'This key is already associated with an element of this collection',
    [VbaErrorCode.FILE_NOT_FOUND]:                   'File not found',
    [VbaErrorCode.BAD_FILE_MODE]:                    'Bad file mode',
    [VbaErrorCode.FILE_ALREADY_OPEN]:                'File already open',
    [VbaErrorCode.FILE_ALREADY_EXISTS]:              'File already exists',
    [VbaErrorCode.TOO_MANY_FILES]:                   'Too many files',
    [VbaErrorCode.PERMISSION_DENIED]:                'Permission denied',
    [VbaErrorCode.PATH_FILE_ACCESS_ERROR]:           'Path/File access error',
    [VbaErrorCode.PATH_NOT_FOUND]:                   'Path not found',
    [VbaErrorCode.OBJECT_VARIABLE_NOT_SET]:          'Object variable or With block variable not set',
    [VbaErrorCode.INVALID_USE_OF_NULL]:              'Invalid use of Null',
    [VbaErrorCode.OBJECT_REQUIRED]:                  'Object required',
    [VbaErrorCode.ARGUMENT_NOT_OPTIONAL]:            'Argument not optional',
    [VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS]:        'Wrong number of arguments or invalid property assignment',
};

/**
 * Throw a plain VBA error object (no line-number enrichment).
 * Use this in non-Evaluator contexts (coerce.ts, vba-types.ts, VbaCollection, etc.).
 * The Evaluator's instance method throwVbaError() adds currentLine and should be
 * preferred inside Evaluator methods.
 */
export function throwVbaError(number: number, message?: string): never {
    throw {
        type: 'VbaError',
        number,
        message: message ?? VBA_ERROR_MESSAGES[number] ?? `Error ${number}`,
    };
}
