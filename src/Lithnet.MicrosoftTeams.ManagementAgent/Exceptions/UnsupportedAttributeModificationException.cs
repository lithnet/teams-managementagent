using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    [Serializable]
    public class UnsupportedAttributeModificationException : TeamsManagementAgentException
    {
        public UnsupportedAttributeModificationException()
        {
        }

        public UnsupportedAttributeModificationException(string message)
            : base(message)
        {
        }

        public UnsupportedAttributeModificationException(string message, Exception inner)
            : base(message, inner)
        {
        }

        protected UnsupportedAttributeModificationException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
