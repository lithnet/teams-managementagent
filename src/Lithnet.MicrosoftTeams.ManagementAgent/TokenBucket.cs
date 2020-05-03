using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NLog;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal class TokenBucket
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        private long availableTokens;

        private readonly object consumerLock;

        private readonly object refillLock;

        private long nextRefillTicks;

        private TimeSpan refillInterval;

        private readonly long refillQuantity;

        private readonly string name;

        public long Capacity { get; }

        public TokenBucket(string name, long capacity, TimeSpan refillInterval, long refillAmount)
        {
            this.name = name;
            this.availableTokens = 0;
            this.consumerLock = new object();
            this.refillLock = new object();
            this.Capacity = capacity;
            this.refillQuantity = refillAmount;
            this.refillInterval = refillInterval;
        }

        public bool TryConsume(CancellationToken token)
        {
            return this.TryConsume(1, token);
        }

        public bool TryConsume(long tokensRequired, CancellationToken token)
        {
            if (tokensRequired <= 0)
            {
                return true;
            }

            if (tokensRequired > this.Capacity)
            {
                throw new ArgumentOutOfRangeException(nameof(tokensRequired), "Number of tokens required is greater than the capacity of the bucket");
            }

            lock (this.consumerLock)
            {
                long newTokens = Math.Min(this.Capacity, this.Refill());

                this.availableTokens = Math.Max(0, Math.Min(this.availableTokens + newTokens, this.Capacity));

                if (tokensRequired > this.availableTokens)
                {
                    return false;
                }
                else
                {
                    this.availableTokens -= tokensRequired;
                    return true;
                }
            }
        }

        public void Consume(CancellationToken token)
        {
            this.Consume(1, token);
        }

        public void Consume(long numTokens, CancellationToken token)
        {
            bool logged = false;
            bool consumed = false;

            if (numTokens <= 0)
            {
                return;
            }

            while (!consumed)
            {
                consumed = this.TryConsume(numTokens, token);

                if (consumed)
                {
                    logger.Trace($"{numTokens} tokens taken from bucket {this.name} leaving {this.availableTokens}");
                }
                else
                {
                    TimeSpan wait = TimeSpan.FromTicks(Math.Max(this.nextRefillTicks - DateTime.Now.Ticks, 1_000_000));
                    if (!logged)
                    {
                        logger.Trace($"{numTokens} tokens not available from bucket {this.name} ({this.availableTokens} remaining). Waiting {wait.TotalMilliseconds} milliseconds");
                        logged = true;
                    }

                    Task.Delay(wait, token).Wait(token);
                }

                token.ThrowIfCancellationRequested();
            }
        }

        public long Refill()
        {
            lock (this.refillLock)
            {
                long now = DateTime.Now.Ticks;

                if (now < this.nextRefillTicks)
                {
                    return 0;
                }

                this.nextRefillTicks = this.refillInterval.Ticks + now;

                logger.Trace($"Refilling bucket {this.name} with {this.refillQuantity} tokens");

                return this.refillQuantity;
            }
        }
    }
}