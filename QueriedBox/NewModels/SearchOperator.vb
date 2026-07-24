Imports System.ComponentModel

Public Enum SearchOperator
    <Description("=")>
    Equal
    <Description("<>")>
    NotEqual
    <Description(">")>
    GreaterThan
    <Description("<")>
    LessThan
    <Description(">=")>
    GreaterThanOrEqual
    <Description("<=")>
    LessThanOrEqual
    <Description("LIKE")>
    [Like]
    <Description("BETWEEN")>
    Between
End Enum