﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal static class AsyncHelper
    {
        private static readonly TaskFactory factory = new TaskFactory(CancellationToken.None, TaskCreationOptions.None, TaskContinuationOptions.None, TaskScheduler.Default);

        // This function destroys ILMerge, so we switched to ILRepack instead. The GetAwaiter() call seems to be responsible for making ILMerge hang
        public static TResult RunSync<TResult>(Task<TResult> func)
        {
            return RunSync(func, CancellationToken.None);
        }

        public static TResult RunSync<TResult>(Task<TResult> func, CancellationToken token)
        {
            return factory.StartNew(() => func, token).Unwrap().GetAwaiter().GetResult();
        }

        public static void RunSync(Task func)
        {
            RunSync(func, CancellationToken.None);
        }

        public static void RunSync(Task func, CancellationToken token)
        {
            factory.StartNew(() => func, token).Unwrap().GetAwaiter().GetResult();
        }

        public static long InterlockedCombine(ref long location,
            Func<long, long> update)
        {
            long initialValue, newValue;

            do
            {
                initialValue = location;
                newValue = update(initialValue);
            }
            while (Interlocked.CompareExchange(ref location, newValue, initialValue) != initialValue);

            return initialValue;
        }

        public static long InterlockedMax(ref long location, long value)
        {
            return InterlockedCombine(ref location,
                v => Math.Max(v, value));
        }
    }
}
