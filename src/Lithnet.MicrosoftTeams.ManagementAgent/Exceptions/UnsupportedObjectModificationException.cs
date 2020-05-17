using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    [Serializable]
    public class UnsupportedObjectModificationException : TeamsManagementAgentException
    {
        public UnsupportedObjectModificationException()
        {
        }

        public UnsupportedObjectModificationException(string message)
            : base(message)
        {
        }

        public UnsupportedObjectModificationException(string message, Exception inner)
            : base(message, inner)
        {
        }

        protected UnsupportedObjectModificationException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
