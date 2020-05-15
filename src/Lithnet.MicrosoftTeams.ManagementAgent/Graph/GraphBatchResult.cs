using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal class GraphBatchResult
    {
        public string ID { get; set; }

        public ErrorResponse ErrorResponse { get; set; }

        public bool IsSuccess { get; set; }

        public bool IsRetryable { get; set; }

        public int RetryInterval { get; set; }

        public bool IsFailed { get; set; }

        public Exception Exception { get; set; }
    }
}