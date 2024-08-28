using System.Reflection;

namespace TesisHelper
{
    internal class StringValueAttribute : Attribute
    {
        public string StringValue { get; private set; }

        public StringValueAttribute(string value)
        {
            StringValue = value;
        }
    }

    public static class EnumExtensions
    {
        public static string GetStringValue(this Enum value)
        {
            FieldInfo field = value.GetType().GetField(value.ToString());
            StringValueAttribute attribute = (StringValueAttribute)field.GetCustomAttribute(typeof(StringValueAttribute));
            return attribute == null ? value.ToString() : attribute.StringValue;
        }
    }
}
