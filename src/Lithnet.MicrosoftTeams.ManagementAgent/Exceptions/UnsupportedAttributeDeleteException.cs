using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    [Serializable]
    public class UnsupportedAttributeDeleteException : UnsupportedAttributeModificationException
    {
        public UnsupportedAttributeDeleteException()
        {
        }

        public UnsupportedAttributeDeleteException(string attributeName)
            : base($"The value of attribute '{attributeName}' can not be deleted")
        {
        }

        public UnsupportedAttributeDeleteException(string message, Exception inner)
            : base(message, inner)
        {
        }

        protected UnsupportedAttributeDeleteException(
            SerializationInfo info,
            StreamingContext context)
            : base(info, context)
        {
        }
    }
}
