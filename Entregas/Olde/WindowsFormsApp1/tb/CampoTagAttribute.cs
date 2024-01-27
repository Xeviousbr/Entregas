using System;

namespace BonifacioEntregas.tb
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
    sealed class CampoTagAttribute : Attribute
    {
        public string Tag { get; }
        public CampoTagAttribute(string tag)
        {
            Tag = tag;
        }
    }

}
