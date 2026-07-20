''' <summary>
''' Attribute used to mark a property that should be ignored
''' when converting a collection to a <see cref="DataTable"/> using Extensions.CollectionExtensions.ToTable"/>.
''' </summary>
''' <remarks>
''' <para>
''' Apply this attribute to any property in a class that represents a data model when you do not
''' want that property to appear as a column in the resulting <see cref="DataTable"/>.
''' </para>
''' <para>
''' Common scenarios include:
''' </para>
''' <list type="bullet">
''' <item>
''' <description>Properties that are calculated at runtime and do not need to be stored in the table.</description>
''' </item>
''' <item>
''' <description>Properties that contain complex objects, references, or sensitive information.</description>
''' </item>
''' <item>
''' <description>Lazy-loaded properties (<c>Lazy(Of T)</c>) that you do not want to evaluate just for table conversion.</description>
''' </item>
''' </list>
''' <para>
''' This attribute ensures that Extensions.CollectionExtensions.ToTable"/> ignores the marked property,
''' preventing runtime exceptions and unnecessary data from appearing in the <see cref="DataTable"/>.
''' </para>
''' </remarks>
<AttributeUsage(AttributeTargets.Property)>
Public Class IgnoreInToTableAttribute
    Inherits Attribute
End Class