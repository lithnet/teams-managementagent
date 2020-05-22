using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    [Serializable]
    public class InvalidProvisioningStateException : TeamsManagementAgentException
    {
        public InvalidProvisioningStateException()
        {
        }

        public InvalidProvisioningStateException(string message)
            : base(message)
        {
        }

        public InvalidProvisioningStateException(string message, Exception inner)
            : base(message, inner)
        {
        }

        protected InvalidProvisioningStateException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
