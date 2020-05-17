using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    [Serializable]
    public class InitialFlowAttributeModificationException : UnsupportedAttributeModificationException
    {
        public InitialFlowAttributeModificationException()
        {
        }

        public InitialFlowAttributeModificationException(string attributeName)
            : base($"This value for attribute '{attributeName}' can not be modified as it can only be set at object creation time as an initial flow attribute")
        {
        }

        public InitialFlowAttributeModificationException(string message, Exception inner)
            : base(message, inner)
        {
        }

        protected InitialFlowAttributeModificationException(
            SerializationInfo info,
            StreamingContext context)
            : base(info, context)
        {
        }
    }
}
