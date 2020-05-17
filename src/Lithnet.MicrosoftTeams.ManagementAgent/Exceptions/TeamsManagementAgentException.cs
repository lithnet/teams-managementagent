using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    [Serializable]
    public class TeamsManagementAgentException : Exception
    {
        public TeamsManagementAgentException()
        {
        }

        public TeamsManagementAgentException(string message)
            : base(message)
        {
        }

        public TeamsManagementAgentException(string message, Exception inner)
            : base(message, inner)
        {
        }

        protected TeamsManagementAgentException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
