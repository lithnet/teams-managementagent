using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Lithnet.Ecma2Framework;
using Microsoft.MetadirectoryServices;

namespace Lithnet.MicrosoftTeams.ManagementAgent.Tests
{
    internal class TestExportContext : IExportContext
    {
        public KeyedCollection<string, ConfigParameter> ConfigParameters { get; } = UnitTestControl.ConfigParameters;

        public CancellationToken Token { get; } = CancellationToken.None;

        public IConnectionContext ConnectionContext { get; } = UnitTestControl.GraphConnectionContext;

        public object CustomData { get; set; }
    }
}
