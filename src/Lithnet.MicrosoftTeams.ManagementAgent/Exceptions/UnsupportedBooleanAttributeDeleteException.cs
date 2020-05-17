using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    [Serializable]
    public class UnsupportedBooleanAttributeDeleteException : UnsupportedAttributeDeleteException
    {
        public UnsupportedBooleanAttributeDeleteException()
        {
        }

        public UnsupportedBooleanAttributeDeleteException(string attributeName)
            : base($"The value of boolean attribute '{attributeName}' can not be deleted. It must be set to true or false", null)
        {
        }

        public UnsupportedBooleanAttributeDeleteException(string message, Exception inner)
            : base(message, inner)
        {
        }

        protected UnsupportedBooleanAttributeDeleteException(
            SerializationInfo info,
            StreamingContext context)
            : base(info, context)
        {
        }
    }
}
