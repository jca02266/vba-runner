# MS-VBAL BNF Grammar

Extracted from `spec/MS-VBAL.txt`.


## §3.2 Module Line Structure

```abnf
module-body-physical-structure = *source-line [non-terminated-line]
source-line = *non-line-termination-character line-terminator
non-terminated-line = *non-line-termination-character
line-terminator = (%x000D  %x000A) / %x000D / %x000A / %x2028 / %x2029
non-line-termination-character = <any character other than %x000D / %x000A / %x2028 / %x2029>
module-body-logical-structure = *extended-line
extended-line = *(line-continuation / non-line-termination-character)  line-terminator
line-continuation = 1*WSC underscore line-terminator
WSC = (tab-character / eom-character /space-character / DBCS-whitespace / most-Unicode-class-Zs)
tab-character = %x0009
eom-character = %x0019
space-character = %x0020
underscore = %x005F
DBCS-whitespace = %x3000
most-Unicode-class-Zs = <all members of Unicode class Zs which are not CP2-characters>
module-body-lines = *logical-line
logical-line = LINE-START *extended-line LINE-END
```

## §3.3.1 Separator and Special Tokens

```abnf
WS = 1*(WSC / line-continuation)
special-token = "," / "." / "!" /  "#" / "&" / "(" / ")" / "*" / "+" / "-" / "/" / ":" / ";" / "<" / "=" / ">" / "?" / "\" / "^"
NO-WS = <no whitespace characters allowed here>
NO-LINE-CONTINUATION = <a line-continuation is not allowed here>
EOL = [WS] LINE-END / single-quote comment-body
EOS = *(EOL  /  ":")  ;End Of Statement
single-quote = %x0027  ; '
comment-body = *(line-continuation / non-line-termination-character) LINE-END
```

## §3.3.2 Number Tokens

```abnf
INTEGER = integer-literal ["%" / "&" / "^"]
integer-literal = decimal-literal / octal-literal / hex-literal
decimal-literal = 1*decimal-digit
octal-literal = "&" [%x004F / %x006F] 1*octal-digit    ; & or &o or &O
hex-literal = "&" (%x0048 / %x0068) 1*hex-digit   ; &h or &H
octal-digit = "0" / "1" / "2" / "3" / "4" / "5" / "6" / "7"
decimal-digit = octal-digit / "8" / "9"
hex-digit = decimal-digit / %x0041-0046 / %x0061-0066 ;A-F / a-f
FLOAT = (floating-point-literal [floating-point-type-suffix] ) / (decimal-literal floating-point-type-suffix)
floating-point-literal = (integer-digits exponent) / (integer-digits "." [fractional-digits] [exponent]) / ( "." fractional-digits [exponent])
integer-digits = decimal-literal
fractional-digits = decimal-literal
exponent = exponent-letter  [sign] decimal-literal
exponent-letter = %x0044 / %x0045 / %x0064 / %x0065   ; D / E / d / e
sign = "+" / "-"
floating-point-type-suffix = "!" / "#" / "@"
```

## §3.3.3 Date Tokens

```abnf
date-or-time = (date-value 1*WSC time-value) / date-value / time-value
date-value = left-date-value date-separator  middle-date-value [date-separator right-date-value]
left-date-value = decimal-literal / month-name
middle-date-value = decimal-literal / month-name
right-date-value = decimal-literal / month-name
date-separator = 1*WSC / (*WSC ("/" / "-" / ",") *WSC)
month-name = English-month-name / English-month-abbreviation
English-month-name = "january" / "february" / "march" / "april" / "may" / "june" / "july" / "august" / "september" / "october" / "november" / "december"
English-month-abbreviation = "jan" / "feb" / "mar" / "apr" / "jun" / "jul" / "aug" / "sep" /  "oct" / "nov" / "dec"
time-value = (hour-value ampm) / (hour-value time-separator minute-value [time-separator second-value] [ampm])
hour-value = decimal-literal
minute-value = decimal-literal
second-value = decimal-literal
time-separator = *WSC (":" / ".") *WSC
ampm = *WSC ("am" / "pm" / "a" / "p")
```

## §3.3.4 String Tokens

```abnf
STRING = double-quote *string-character (double-quote /  line-continuation / LINE-END)
double-quote = %x0022  ; "
string-character = NO-LINE-CONTINUATION ((double-quote double-quote)  /  non-line-termination-character)
```

## §3.3.5 Identifier Tokens

```abnf
lex-identifier = Latin-identifier / codepage-identifier / Japanese-identifier / Korean-identifier / simplified-Chinese-identifier / traditional-Chinese-identifier
Latin-identifier = first-Latin-identifier-character *subsequent-Latin-identifier-character
first-Latin-identifier-character = (%x0041-005A / %x0061-007A) ; A-Z / a-z
subsequent-Latin-identifier-character = first-Latin-identifier-character / decimal-digit / %x5F    ; underscore
```

## §3.3.5.1 Non-Latin Identifiers

```abnf
Japanese-identifier = first-Japanese-identifier-character *subsequent-Japanese-identifier-character
first-Japanese-identifier-character = (first-Latin-identifier-character / CP932-initial-character)
subsequent-Japanese-identifier-character = (subsequent-Latin-identifier-character / CP932-subsequent-character)
CP932-initial-character = < character ranges specified in section 3.3.5.1.1>
CP932-subsequent-character = < character ranges specified in section 3.3.5.1.1>
Korean-identifier = first-Korean-identifier-character *subsequent Korean-identifier-character
first-Korean-identifier-character = (first-Latin-identifier-character / CP949-initial-character )
subsequent-Korean-identifier-character = (subsequent-Latin-identifier-character / CP949-subsequent-character)
CP949-initial-character = < character ranges specified in section 3.3.5.1.2>
CP949-subsequent-character = < character ranges specified in section 3.3.5.1.2>
simplified-Chinese-identifier = first-sChinese-identifier-character
first-sChinese-identifier-character = (first-Latin-identifier-character / CP936-initial-character)
subsequent-sChinese-identifier-character = (subsequent-Latin-identifier-character / CP936-subsequent-character)
CP936-initial-character = < character ranges specified in section 3.3.5.1.3>
CP936-subsequent-character = < character ranges specified in section 3.3.5.1.3>
traditional-Chinese-identifier = first-tChinese-identifier-character
first-tChinese-identifier-character = (first-Latin-identifier-character / CP950-initial-character)
subsequent-tChinese-identifier-character = (subsequent-Latin-identifier-character / CP950-subsequent-character)
CP950-initial-character = < character ranges specified in section 3.3.5.1.4>
CP950-subsequent-character = < character ranges specified in section 3.3.5.1.4>
codepage-identifier = (first-Latin-identifier-character / CP2-character)
CP2-character = <any Unicode character that has a mapping to the character range %x80-FF in a Microsoft Windows supported code page>
```

## §3.3.5.2 Reserved Identifiers and IDENTIFIER

```abnf
reserved-identifier = statement-keyword / marker-keyword / operator-identifier /
IDENTIFIER = <any lex-identifier that is not a reserved-identifier>
statement-keyword = "Call" / "Case" /"Close" / "Const"/ "Declare" / "DefBool" / "DefByte" / "DefCur" / "DefDate" / "DefDbl" / "DefInt" / "DefLng" / "DefLngLng" / "DefLngPtr" / "DefObj" / "DefSng" / "DefStr" / "DefVar" / "Dim" / "Do" / "Else" / "ElseIf" / "End" / "EndIf" /  "Enum" / "Erase" / "Event" / "Exit" / "For" / "Friend" / "Function" / "Get" / "Global" / "GoSub" / "GoTo" / "If" / "Implements"/ "Input" / "Let" / "Lock" / "Loop" / "LSet" / "Next" / "On" / "Open" / "Option" / "Print" / "Private" / "Public" / "Put" / "RaiseEvent" / "ReDim" / "Resume" / "Return" / "RSet" / "Seek" / "Select" / "Set" / "Static" / "Stop" / "Sub" / "Type" / "Unlock" / "Wend" / "While" / "With" / "Write"
rem-keyword = "Rem"
marker-keyword = "Any" / "As"/ "ByRef" / "ByVal "/"Case" / "Each" / "Else" /"In"/ "New" / "Shared" / "Until" / "WithEvents" / "Write" / "Optional" / "ParamArray" / "Preserve" / "Spc" / "Tab" / "Then" / "To"
operator-identifier = "AddressOf" / "And" / "Eqv" / "Imp" / "Is" / "Like" / "New" / "Mod" / "Not" / "Or" / "TypeOf" / "Xor"
```

## §6.1.2.8.1.5 DoEvents

```abnf
reserved-name = "Abs" / "CBool" / "CByte" / "CCur" / "CDate" / "CDbl" / "CDec" / "CInt" / "CLng" / "CLngLng" / "CLngPtr" / "CSng" / "CStr" / "CVar" / "CVErr" / "Date" / "Debug" / "DoEvents" / "Fix" / "Int" / "Len" / "LenB" / "Me" / "PSet" / "Scale" / "Sgn" / "String"
special-form = "Array" / "Circle" / "Input" / "InputB"  / "LBound" / "Scale" / "UBound"
reserved-type-identifier = "Boolean" / "Byte" / "Currency" / "Date" / "Double" /  "Integer" / "Long" / "LongLong" / "LongPtr" / "Single" / "String" / "Variant"
literal-identifier = boolean-literal-identifier / object-literal-identifier / variant-literal-identifier
boolean-literal-identifier = "true" / "false"
object-literal-identifier = "nothing"
variant-literal-identifier = "empty" / "null"
```

## §6.1.3.2.2.1 Description

```abnf
reserved-for-implementation-use = "Attribute" / "LINEINPUT" / "VB_Base" / "VB_Control" / "VB_Creatable" /  "VB_Customizable" / "VB_Description" / "VB_Exposed" / "VB_Ext_KEY " / "VB_GlobalNameSpace" / "VB_HelpID" / "VB_Invoke_Func" / "VB_Invoke_Property " / "VB_Invoke_PropertyPut" / "VB_Invoke_PropertyPutRef" / "VB_MemberFlags" / "VB_Name" / "VB_PredeclaredId" / "VB_ProcData" / "VB_TemplateDerived" / "VB_UserMemId" / "VB_VarDescription" / "VB_VarHelpID" / "VB_VarMemberFlags" / "VB_VarProcData " / "VB_VarUserMemId"
future-reserved = "CDecl" / "Decimal" / "DefDec"
```

## §3.3.5.3 Special Identifier Forms

```abnf
FOREIGN-NAME = "[" foreign-identifier "]"
foreign-identifier = 1*non-line-termination-character
BUILTIN-TYPE = reserved-type-identifier /  ("[" reserved-type-identifier "]") / "object" / "[object]"
TYPED-NAME = IDENTIFIER  type-suffix
type-suffix = "%" / "&" / "^" / "!" / "#" / "@" / "$"
conditional-module-body = cc-block
cc-block = *(cc-const / cc-if-block / logical-line)
```

## §3.4.1 Conditional Compilation Const Directive

```abnf
cc-const = LINE-START  "#"  "const" cc-var-lhs "=" cc-expression cc-eol
cc-var-lhs = name
cc-eol = [single-quote *non-line-termination-character] LINE-END
```

## §3.4.2 Conditional Compilation If Directives

```abnf
cc-if-block = cc-if
cc-if = LINE-START  "#" "if" cc-expression "then" cc-eol
cc-elseif-block = cc-elseif cc-block
cc-elseif = LINE-START "#" "elseif" cc-expression "then" cc-eol
cc-else-block = cc-else cc-block
cc-else = LINE-START "#" "else" cc-eol
cc-endif = LINE-START "#" ("endif" / ("end" "if")) cc-eol
```

## §4.2 Modules

```abnf
procedural-module = LINE-START procedural-module-header EOS
```

## §5.1 Module Body Structure

```abnf
procedural-module-body = LINE-START  procedural-module-declaration-section
class-module-body = LINE-START  class-module-declaration-section
unrestricted-name = name / reserved-identifier
name = untyped-name / TYPED-NAME / file-number
untyped-name = IDENTIFIER / FOREIGN-NAME
```

## §5.2 Module Declaration Section Structure

```abnf
procedural-module-declaration-section = [*(procedural-module-directive-element EOS) def-directive]  *( procedural-module-declaration-element EOS)
class-module-declaration-section = [*(class-module-directive-element EOS) def-directive]  *(class-module-declaration-element EOS)
procedural-module-directive-element = common-option-directive / option-private-directive / def-directive
procedural-module-declaration-element = common-module-declaration-element / global-variable-declaration / public-const-declaration / public-type-declaration / public-external-procedure-declaration / global-enum-declaration / common-option-directive / option-private-directive
class-module-directive-element = common-option-directive / def-directive / implements-directive
class-module-declaration-element = common-module-declaration-element / event-declaration / commonoption-directive / implements-directive
```

## §5.2.1 Option Directives

```abnf
common-option-directive = option-compare-directive /  option-base-directive / option-explicit-directive  / rem-statement
```

## §5.2.1.1 Option Compare Directive

```abnf
option-compare-directive = "Option"   "Compare"   ( "Binary" / "Text")
```

## §5.2.1.2 Option Base Directive

```abnf
option-base-directive = "Option"   "Base"     INTEGER
```

## §5.2.1.3 Option Explicit Directive

```abnf
option-explicit-directive = "Option"   "Explicit"
```

## §5.2.1.4 Option Private Directive

```abnf
option-private-directive = "Option"   "Private"   "Module"
```

## §5.2.2 Implicit Definition Directives

```abnf
def-directive = def-type  letter-spec *( "," letter-spec)
letter-spec = single-letter /  universal-letter-range / letter-range
single-letter = IDENTIFIER   ; %x0041-005A / %x0061-007A
universal-letter-range = upper-case-A "-"upper-case-Z
upper-case-A = IDENTIFIER
upper-case-Z = IDENTIFIER
letter-range = first-letter  "-" last-letter
first-letter = IDENTIFIER
last-letter = IDENTIFIER
def-type = "DefBool" / "DefByte" / "DefCur" /  "DefDate" / "DefDbl" / "DefInt" / "DefLng" / "DefLngLng" / "DefLngPtr" / "DefObj" / "DefSng" / "DefStr" / "DefVar"
```

## §5.2.3 Module Declarations

```abnf
common-module-declaration-element = module-variable-declaration
common-module-declaration-element =/ private-const-declaration
common-module-declaration-element =/ private-type-declaration
common-module-declaration-element =/ public-type-declaration
common-module-declaration-element =/ public-enum-declaration
common-module-declaration-element =/ private-enum-declaration
common-module-declaration-element =/ private-external-procedure-declaration
common-module-declaration-element =/ attribute-statement
```

## §5.2.3.1 Module Variable Declaration Lists

```abnf
module-variable-declaration = public-variable-declaration / private-variable-declaration
global-variable-declaration = "Global"  variable-declaration-list
public-variable-declaration = "Public" ["Shared"] module-variable-declaration-list
private-variable-declaration = ("Private" / "Dim") [ "Shared"] module-variable-declaration-list
module-variable-declaration-list = (withevents-variable-dcl / variable-dcl)
variable-declaration-list = variable-dcl *( "," variable-dcl )
```

## §5.2.3.1.1 Variable Declarations

```abnf
variable-dcl = typed-variable-dcl / untyped-variable-dcl
typed-variable-dcl = TYPED-NAME [array-dim]
untyped-variable-dcl = IDENTIFIER  [array-clause / as-clause]
array-clause = array-dim [as-clause]
as-clause = as-auto-object / as-type
```

## §5.2.3.1.2 WithEvents Variable Declarations

```abnf
withevents-variable-dcl = "withevents" IDENTIFIER "as" class-type-name
class-type-name = defined-type-expression
```

## §5.2.3.1.3 Array Dimensions and Bounds

```abnf
array-dim = "(" [bounds-list] ")"
bounds-list = dim-spec *("," dim-spec)
dim-spec = [lower-bound] upper-bound
lower-bound = constant-expression  "to"
upper-bound = constant-expression
```

## §5.2.3.1.4 Variable Type Declarations

```abnf
as-auto-object = "as" "new" class-type-name
as-type = "as" type-spec
type-spec = fixed-length-string-spec  / type-expression
fixed-length-string-spec = "string" "*" string-length
string-length = constant-name / INTEGER
constant-name = simple-name-expression
```

## §5.2.3.2 Const Declarations

```abnf
public-const-declaration = ("Global" / "Public")  module-const-declaration
private-const-declaration = ["Private"] module-const-declaration
module-const-declaration = const-declaration
const-declaration = "Const"  const-item-list
const-item-list = const-item *[ "," const-item]
const-item = typed-name-const-item / untyped-name-const-item
typed-name-const-item = TYPED-NAME "=" constant-expression
untyped-name-const-item = IDENTIFIER [const-as-clause] "=" constant-expression
const-as-clause = "as" BUILTIN-TYPE
```

## §5.2.3.3 User Defined Type Declarations

```abnf
public-type-declaration = ["global" / "public"]  udt-declaration
private-type-declaration = "private" udt-declaration
udt-declaration = "type" untyped-name EOS udt-member-list EOS "end" "type"
udt-member-list = udt-element *[EOS udt-element]
udt-element = rem-statement / udt-member
udt-member = reserved-name-member-dcl / untyped-name-member-dcl
untyped-name-member-dcl = IDENTIFIER optional-array-clause
reserved-name-member-dcl = reserved-member-name as-clause
optional-array-clause = [array-dim] as-clause
reserved-member-name = statement-keyword / marker-keyword / operator-identifier / special-form / reserved-name / literal-identifier / reserved-for-implementation-use / future-reserved
```

## §5.2.3.4 Enum Declarations

```abnf
global-enum-declaration = "global" enum-declaration
public-enum-declaration = ["public"] enum-declaration
private-enum-declaration = "private" enum-declaration
enum-declaration = "enum" untyped-name EOS enum-member-list EOS "end" "enum"
enum-member-list = enum-element *[EOS enum-element]
enum-element = rem-statement / enum-member
enum-member = untyped-name [ "=" constant-expression]
```

## §5.2.3.5 External Procedure Declaration

```abnf
public-external-procedure-declaration = ["public"] external-proc-dcl
private-external-procedure-declaration = "private" external-proc-dcl
external-proc-dcl = "declare" ["ptrsafe"] (external-sub / external-function)
external-sub = "sub" subroutine-name lib-info [procedure-parameters]
external-function = "function" function-name lib-info  [procedure-parameters] [function-type]
lib-info = lib-clause [alias-clause]
lib-clause = "lib" STRING
alias-clause = "alias" STRING
```

## §5.2.4.1.2 Default Instance Variables Static Semantics

```abnf
implements-directive = "Implements" class-type-name
event-declaration = ["Public"]
event-parameter-list = "(" [positional-parameters] ")"
```

## §5.3 Module Code Section Structure

```abnf
procedural-module-code-section = *( LINE-START  procedural-module-code-element LINE-END)
class-module-code-section = *( LINE-START  class-module-code-element LINE-END)
procedural-module-code-element = common-module-code-element
class-module-code-element = common-module-code-element / implements-directive
common-module-code-element = rem-statement / procedure-declaration
procedure-declaration = subroutine-declaration / function-declaration /        property-get-declaration / property-LHS-declaration
```

## §5.3.1 Procedure Declarations

```abnf
subroutine-declaration = procedure-scope [initial-static]
function-declaration = procedure-scope [initial-static]
property-get-declaration = procedure-scope [initial-static]
property-lhs-declaration = procedure-scope [initial-static]
end-label = statement-label-definition
procedure-tail = [WS] LINE-END / single-quote comment-body /  ":" rem-statement
```

## §5.3.1.1 Procedure Scope

```abnf
procedure-scope = ["global" / "public" / "private" / "friend"]
```

## §5.3.1.2 Static Procedures

```abnf
initial-static = "static"
trailing-static = "static"
```

## §5.3.1.3 Procedure Names

```abnf
subroutine-name = IDENTIFIER / prefixed-name
function-name = TYPED-NAME / subroutine-name
prefixed-name = event-handler-name / implemented-name / lifecycle-handler-name
```

## §5.3.1.4 Function Type Declarations

```abnf
function-type = "as" type-expression [array-designator]
array-designator = "(" ")"
```

## §5.3.1.5 Parameter Lists

```abnf
procedure-parameters = "(" [parameter-list] ")"
property-parameters = "(" [parameter-list ","] value-param ")"
parameter-list = (positional-parameters "," optional-parameters ) /
positional-parameters = positional-param *("," positional-param)
optional-parameters = optional-param *("," optional-param)
value-param = positional-param
positional-param = [parameter-mechanism] param-dcl
optional-param = optional-prefix param-dcl [default-value]
param-array = "paramarray" IDENTIFIER "(" ")" ["as" ("variant" / "[variant]")]
param-dcl = untyped-name-param-dcl / typed-name-param-dcl
untyped-name-param-dcl = IDENTIFIER [parameter-type]
typed-name-param-dcl = TYPED-NAME [array-designator]
optional-prefix = ("optional" [parameter-mechanism]) / ([parameter-mechanism] ("optional"))
parameter-mechanism = "byval" / " byref"
parameter-type = [array-designator] "as" (type-expression / "Any")
default-value = "=" constant-expression
```

## §5.3.1.8 Event Handler Declarations

```abnf
event-handler-name = IDENTIFIER
```

## §5.3.1.9 Implemented Name Declarations

```abnf
implemented-name = IDENTIFIER
```

## §5.3.1.10 Lifecycle Handler Declarations

```abnf
lifecycle-handler-name = “Class_Initialize” / “Class_Terminate”
```

## §5.4 Procedure Bodies and Statements

```abnf
procedure-body = statement-block
```

## §5.4.1 Statement Blocks

```abnf
statement-block = *(block-statement EOS)
block-statement = statement-label-definition / rem-statement / statement / attribute-statement
attribute-statement = attribute [IDENTIFIER "."] reserved-for-implementation-use attr-eq [quoted-identifier / boolean-literal-identifier]
statement = control-statement / data-manipulation-statement / error-handling-statement / file-statement
```

## §5.4.1.1 Statement Labels

```abnf
statement-label-definition = LINE-START ((identifier-statement-label “:”) / (line-number-label [“:”] ))
statement-label = identifier-statement-label / line-number-label
statement-label-list = statement-label [“,” statement-label]
identifier-statement-label = IDENTIFIER
line-number-label = INTEGER
```

## §5.4.1.2 Rem Statement

```abnf
rem-statement = "Rem" comment-body
```

## §5.4.2 Control Statements

```abnf
control-statement = if-statement / control-statement-except-multiline-if
control-statement-except-multiline-if = call-statement / while-statement / for-statement / exit-for-statement / do-statement / exit-do-statement / single-line-if-statement /  select-case-statement /stop-statement / goto-statement / on-goto-statement / gosub-statement / return-statement / on-gosub-statement /for-each-statement / exit-sub-statement / exit-function-statement / exit-property-statement / raiseevent-statement / with-statement / end-statement / assert-statement
```

## §5.4.2.1 Call Statement

```abnf
call-statement = "Call" (simple-name-expression / member-access-expression / index-expression / with-expression)
call-statement =/ (simple-name-expression / member-access-expression / with-expression) argument-list
```

## §5.4.2.2 While Statement

```abnf
while-statement = "While" boolean-expression EOS  statement-block  "Wend"
```

## §5.4.2.3 For Statement

```abnf
for-statement = simple-for-statement / explicit-for-statement
simple-for-statement = for-clause EOS statement-block “Next”
explicit-for-statement = for-clause EOS statement-block
nested-for-statement = explicit-for-statement / explicit-for-each-statement
for-clause = “For” bound-variable-expression “=” start-value “To” end-value [step-clause]
start-value = expression
end-value = expression
step-clause = Step" step-increment
step-increment = expression
```

## §5.4.2.4 For Each Statement

```abnf
for-each-statement = simple-for-each-statement / explicit-for-each-statement
simple-for-each-statement = for-each-clause EOS statement-block “Next”
explicit-for-each-statement = for-each-clause EOS statement-block
for-each-clause = “For” “Each” bound-variable-expression “In” collection
collection = expression
```

## §5.4.2.5 Exit For Statement

```abnf
exit-for-statement = "Exit" "For"
```

## §5.4.2.6 Do Statement

```abnf
do-statement = "Do" [condition-clause] EOS statement-block
condition-clause = while-clause / until-clause
while-clause = "While" boolean-expression
until-clause = "Until" boolean-expression
```

## §5.4.2.7 Exit Do Statement

```abnf
exit-do-statement = "Exit" "Do"
```

## §5.4.2.8 If Statement

```abnf
if-statement = LINE-START "If" boolean-expression "Then" EOL statement-block
else-if-block = LINE-START "ElseIf" boolean-expression "Then" EOL
else-if-block =/ "ElseIf" boolean-expression "Then" statement-block
else-block = LINE-START "Else" statement-block
```

## §5.4.2.9 Single-line If Statement

```abnf
single-line-if-statement = if-with-non-empty-then / if-with-empty-then
if-with-non-empty-then = "If" boolean-expression "Then" list-or-label [single-line-else-clause]
if-with-empty-then = "If" boolean-expression "Then" single-line-else-clause
single-line-else-clause = "Else" [list-or-label]
list-or-label = (statement-label *[":" [same-line-statement]]) /
same-line-statement = file-statement / error-handling-statement /
```

## §5.4.2.10 Select Case Statement

```abnf
select-case-statement = "Select" "Case" WS select-expression EOS
case-clause = "Case" range-clause *("," range-clause) EOS statement-block
case-else-clause = "Case" "Else" EOS statement-block
range-clause = expression
range-clause =/  start-value "To" end-value
range-clause =/    ["Is"] comparison-operator expression
start-value = expression
end-value = expression
select-expression = expression
comparison-operator = "=" / ("<" ">" ) / (">" "<") / "<" / ">" / (">" "=") /  ("=" ">") / ("<" "=") / ("=" "<")
```

## §5.4.2.11 Stop Statement

```abnf
stop-statement = "Stop"
```

## §5.4.2.12 GoTo Statement

```abnf
goto-statement = (("Go" "To") / "GoTo") statement-label
```

## §5.4.2.13 On…GoTo Statement

```abnf
on-goto-statement = "On" expression "GoTo" statement-label-list
```

## §5.4.2.14 GoSub Statement

```abnf
gosub-statement = (("Go" "Sub") / "GoSub") statement-label
```

## §5.4.2.15 Return Statement

```abnf
return-statement = "Return"
```

## §5.4.2.16 On…GoSub Statement

```abnf
on-gosub-statement = "On" expression "GoSub" statement-label-list
```

## §5.4.2.17 Exit Sub Statement

```abnf
exit-sub-statement = "Exit" "Sub"
```

## §5.4.2.18 Exit Function Statement

```abnf
exit-function-statement = "Exit" "Function"
```

## §5.4.2.19 Exit Property Statement

```abnf
exit-property-statement = "Exit" "Property"
```

## §5.4.2.20 RaiseEvent Statement

```abnf
raiseevent-statement = "RaiseEvent" IDENTIFIER ["(" event-argument-list ")"]
event-argument-list = [event-argument *("," event-argument)]
event-argument = expression
```

## §5.4.2.21 With Statement

```abnf
with-statement = "With" expression EOS statement-block "End" "With"
```

## §5.4.2.22 End Statement

```abnf
end-statement = "End"
```

## §5.4.2.23 Assert Statement

```abnf
assert-statement = "Debug" "." "Assert" boolean-expression
```

## §5.4.3 Data Manipulation Statements

```abnf
data-manipulation-statement = local-variable-declaration / static-variable-declaration / local-const-declaration / redim-statement / erase-statement / mid-statement /rset-statement / lset-statement / let-statement / set-statement
```

## §5.4.3.1 Local Variable Declarations

```abnf
local-variable-declaration = ("Dim" ["Shared"] variable-declaration-list)
static-variable-declaration = "Static" variable-declaration-list
```

## §5.4.3.2 Local Constant Declarations

```abnf
local-const-declaration = const-declaration
```

## §5.4.3.3 ReDim Statement

```abnf
redim-statement = "Redim" ["Preserve"] redim-declaration-list
redim-declaration-list = redim-variable-dcl *("," redim-variable-dcl)
redim-variable-dcl = redim-typed-variable-dcl / redim-untyped-dcl / with-expression-dcl / member-access-expression-dcl
redim-typed-variable-dcl = TYPED-NAME dynamic-array-dim
redim-untyped-dcl = untyped-name dynamic-array-clause
with-expression-dcl = with-expression dynamic-array-clause
member-access-expression-dcl = member-access-expression dynamic-array-clause
dynamic-array-dim = "(" dynamic-bounds-list ")"
dynamic-bounds-list = dynamic-dim-spec *[ "," dynamic-dim-spec ]
dynamic-dim-spec = [dynamic-lower-bound] dynamic-upper-bound
dynamic-lower-bound = integer-expression  "to"
dynamic-upper-bound = integer-expression
dynamic-array-clause = dynamic-array-dim [as-clause]
```

## §5.4.3.4 Erase Statement

```abnf
erase-statement = “Erase” erase-list
erase-list = erase-element *[ “,” erase-element]
erase-element = l-expression
```

## §6.1.2.11.1.27 MidB$

```abnf
mid-statement = mode-specifier "(" string-argument "," start ["," length] ")" "=" expression
mode-specifier = ("Mid" / "MidB" / "Mid$" / "MidB$")
string-argument = bound-variable-expression
start = integer-expression
length = integer-expression
```

## §5.4.3.9 Set Statement

```abnf
lset-statement = "LSet" bound-variable-expression "=" expression
```

## §5.4.3.7 RSet Statement

```abnf
rset-statement = "RSet" bound-variable-expression "=" expression
```

## §5.4.3.8 Let Statement

```abnf
let-statement = ["Let"] l-expression "=" expression
set-statement = "Set" l-expression "=" expression
```

## §5.4.4 Error Handling Statements

```abnf
error-handling-statement = on-error-statement / resume-statement / error-statement
```

## §5.4.4.3 Error Statement

```abnf
on-error-statement = "On" "Error" error-behavior
error-behavior = ("Resume" "Next") / ("GoTo" (statement-label / -1))
```

## §5.4.4.2 Resume Statement

```abnf
resume-statement = "Resume" [("Next" / statement-label)]
Error-statement = "Error" error-number
error-number = integer-expression
```

## §5.4.5 File Statements

```abnf
file-statement = open-statement / close-statement / seek-statement / lock-statement / unlock-statement / line-input-statement / width-statement / print-statement / write-statement / input-statement / put-statement / get-statement
```

## §5.4.5.1 Open Statement

```abnf
open-statement = "Open" path-name [mode-clause] [access-clause] [lock] "As" file-number [len-clause]
path-name = expression
mode-clause = "For" mode
mode = "Append" / "Binary" / "Input" / "Output" / "Random"
access-clause = "Access" access
access = "Read" / "Write" / ("Read" "Write")
lock = "Shared" / ("Lock"  "Read") / ("Lock" "Write") / ("Lock" "Read" "Write")
len-clause = "Len" "=" rec-length
rec-length = expression
```

## §5.4.5.1.1 File Numbers

```abnf
file-number = marked-file-number / unmarked-file-number
marked-file-number = "#" expression
unmarked-file-number = expression
```

## §5.4.5.2 Close and Reset Statements

```abnf
close-statement = "Reset" / ("Close" [file-number-list])
file-number-list = file-number *[ "," file-number]
```

## §5.4.5.3 Seek Statement

```abnf
seek-statement = "Seek" file-number "," position
position = expression
```

## §5.4.5.4 Lock Statement

```abnf
lock-statement = "Lock" file-number [ "," record-range]
record-range = start-record-number / ([start-record-number] "To" end-record-number)
start-record-number = expression
end-record-number = expression
```

## §5.4.5.5 Unlock Statement

```abnf
unlock-statement = "Unlock" file-number [ "," record-range]
```

## §5.4.5.10 Input Statement

```abnf
line-input-statement = "Line"  "Input" marked-file-number "," variable-name
variable-name = variable-expression
```

## §5.4.5.7 Width Statement

```abnf
width-statement = "Width"   marked-file-number   ","  line-width
line-width = expression
```

## §5.4.5.8 Print Statement

```abnf
print-statement = [("Debug" / "Me") "."] "Print" marked-file-number "," [output-list]
```

## §5.4.5.8.1 Output Lists

```abnf
output-list = *output-item
output-item = [output-clause] [char-position]
output-clause = (spc-clause / tab-clause / output-expression)
char-position = ( ";" / ",")
output-expression = expression
spc-clause = "Spc" "(" spc-number ")"
spc-number = expression
tab-clause = "Tab" [tab-number-clause]
tab-number-clause = "(" tab-number ")"
tab-number = expression
```

## §5.4.5.9 Write Statement

```abnf
write-statement = "Write" marked-file-number "," [output-list]
input-statement = "Input" marked-file-number "," input-list
input-list = input-variable *[ ","  input-variable]
input-variable = bound-variable-expression
```

## §5.4.5.11 Put Statement

```abnf
put-statement = "Put" file-number ","[record-number] "," data
record-number = expression
data = expression
```

## §5.4.5.12 Get Statement

```abnf
get-statement = "Get" file-number "," [record-number] "," variable
variable = variable-expression
```

## §5.5.1.2.4 Let-coercion to and from String

```abnf
numeric-coercion-string = [WS] [sign [WS]] regionalnumber-string [exponentclause] [WS]
exponent-clause = ["e" / "d"] [sign] integer-literal
sign = "+" / "-"
regional-number-string = <unsigned number or currency value interpreted according to the active host-defined regional settings>
```

## §5.5.2.2.2 Set-coercion to and from non-object types

```abnf
expression = value-expression / l-expression
value-expression = literal-expression / parenthesized-expression / typeof-is-expression / new-expression / operator-expression
l-expression = simple-name-expression / instance-expression / member-access-expression / index-expression / dictionary-access-expression / with-expression
```

## §5.6.5 Literal Expressions

```abnf
literal-expression = INTEGER / FLOAT / DATE / STRING / (literal-identifier [type-suffix])
```

## §5.6.6 Parenthesized Expressions

```abnf
parenthesized-expression = "(" expression ")"
```

## §5.6.7 TypeOf…Is Expressions

```abnf
typeof-is-expression = "typeof" expression "is" type-expression
```

## §5.6.8 New Expressions

```abnf
new-expression = "New" type-expression
```

## §5.6.9 Operator Expressions

```abnf
operator-expression = arithmetic-operator-expression / concatenation-operator-expression / relational-operator-expression / like-operator-expression / is-operator-expression / logical-operator-expression
```

## §5.6.9.3 Arithmetic Operators

```abnf
arithmetic-operator-expression = unary-minus-operator-expression / addition-operator-expression / subtraction-operator-expression / multiplication-operator-expression / division-operator-expression / integer-division-operator-expression / modulo-operator-expression / exponentiation-operator-expression
```

## §5.6.9.3.1 Unary - Operator

```abnf
unary-minus-operator-expression = "-" expression
addition-operator-expression = expression "+" expression
```

## §5.6.9.3.3 Binary - Operator

```abnf
subtraction-operator-expression = expression "-" expression
multiplication-operator-expression = expression "*" expression
division-operator-expression = expression "/" expression
integer-division-operator-expression = expression "\" expression
modulo-operator-expression = expression "mod" expression
exponentiation-operator-expression = expression "^" expression
concatenation-operator-expression = expression "&" expression
```

## §5.6.9.5 Relational Operators

```abnf
relational-operator-expression = equality-operator-expression / inequality-operator-expression / less-than-operator-expression / greater-than-operator-expression / less-than-equal-operator-expression / greater-than-equal-operator-expression
The = operator performs a value equality comparison on its operands.
equality-operator-expression = expression "=" expression
inequality-operator-expression = expression ( "<"">" / ">""<" ) expression
less-than-operator-expression = expression "<" expression
greater-than-operator-expression = expression ">" expression
less-than-equal-operator-expression = expression ( "<""=" / "=""<" ) expression
greater-than-equal-operator-expression = expression ( ">""=" / "="">" ) expression
```

## §5.6.9.6 Like Operator

```abnf
like-operator-expression = expression "like" like-pattern-expression
like-pattern-expression = expression
like-pattern-string = *like-pattern-element
like-pattern-element = like-pattern-char / "?" / "#" / "*" / like-pattern-charlist
like-pattern-char = <Any character except "?", "#", "*" and "[" >
like-pattern-charlist = "[" ["!"] ["-"] *like-pattern-charlist-element ["-"] "]"
like-pattern-charlist-element = like-pattern-charlist-char / like-pattern-charlist-range
like-pattern-charlist-range = like-pattern-charlist-char "-" like-pattern-charlist-char
like-pattern-charlist-char = <Any character except "-" and "]">
```

## §5.6.9.7 Is Operator

```abnf
is-operator-expression = expression "is" expression
```

## §5.6.9.8 Logical Operators

```abnf
logical-operator-expression = not-operator-expression / and-operator-expression / or-operator-expression / xor-operator-expression / imp-operator-expression / eqv-operator-expression
```

## §5.6.9.8.1 Not Operator

```abnf
not-operator-expression = "not" expression
```

## §5.6.9.8.2 And Operator

```abnf
and-operator-expression = expression "and" expression
```

## §5.6.9.8.3 Or Operator

```abnf
or-operator-expression = expression "or" expression
```

## §5.6.9.8.4 Xor Operator

```abnf
xor-operator-expression = expression "xor" expression
```

## §5.6.9.8.5 Eqv Operator

```abnf
eqv-operator-expression = expression "eqv" expression
```

## §5.6.9.8.6 Imp Operator

```abnf
imp-operator-expression = expression "imp" expression
simple-name-expression = name / special-form / reserved-name
```

## §5.6.11 Instance Expressions

```abnf
instance-expression = "me"
```

## §5.6.12 Member Access Expressions

```abnf
member-access-expression = l-expression NO-WS "." unrestricted-name
member-access-expression =/ l-expression line-continuation "." unrestricted-name
```

## §8 Index

```abnf
index-expression = l-expression "(" argument-list ")"
```

## §5.6.13.1 Argument Lists

```abnf
argument-list = [positional-or-named-argument-list]
positional-or-named-argument-list = *(positional-argument ",") required-positional-argument
positional-or-named-argument-list =/   *(positional-argument ",") named-argument-list
positional-argument = [argument-expression]
required-positional-argument = argument-expression
named-argument-list = named-argument *("," named-argument)
named-argument = unrestricted-name ":""=" argument-expression
argument-expression = ["byval"] expression
argument-expression =/  addressof-expression
```

## §5.6.14 Dictionary Access Expressions

```abnf
dictionary-access-expression = l-expression  NO-WS "!" NO-WS unrestricted-name
dictionary-access-expression =/  l-expression  line-continuation "!" NO-WS unrestricted-name
dictionary-access-expression =/  l-expression  line-continuation "!" line-continuation unrestricted-name
```

## §5.6.15 With Expressions

```abnf
with-expression = with-member-access-expression / with-dictionary-access-expression
with-member-access-expression = "." unrestricted-name
with-dictionary-access-expression = "!" unrestricted-name
```

## §5.6.16.1 Constant Expressions

```abnf
constant-expression = expression
```

## §5.6.16.2 Conditional Compilation Expressions

```abnf
cc-expression = expression
```

## §5.6.16.3 Boolean Expressions

```abnf
boolean-expression = expression
```

## §5.6.16.4 Integer Expressions

```abnf
integer-expression = expression
```

## §5.6.16.5 Variable Expressions

```abnf
variable-expression = l-expression
```

## §5.6.16.6 Bound Variable Expressions

```abnf
bound-variable-expression = l-expression
```

## §5.6.16.7 Type Expressions

```abnf
type-expression = BUILTIN-TYPE / defined-type-expression
defined-type-expression = simple-name-expression / member-access-expression
```

## §5.6.16.8 AddressOf Expressions

```abnf
addressof-expression = "addressof" procedure-pointer-expression
procedure-pointer-expression = simple-name-expression / member-access-expression
```

## §6.1.3.2.2.3 HelpFile

```abnf
VbMsgBoxStyle = vbOKOnly, Optional Title As Variant,
```

## §6.1.2.11.1.7 Filter

```abnf
VbCompareMethod = vbBinaryCompare)
```

## §6.1.2.11.1.14 InStr / InStrB

```abnf
VbCompareMethod = vbBinaryCompare)
```
