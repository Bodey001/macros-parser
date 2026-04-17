namespace VbaMacroParser.Models;

public enum ModuleType
{
    Standard,
    Class,
    Form
}

public enum ProcedureKind
{
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet
}

public enum AccessModifier
{
    Default,
    Public,
    Private,
    Friend
}
