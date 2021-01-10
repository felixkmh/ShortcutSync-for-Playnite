using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ShortcutSync
{
    class TaskQueue
    {
        protected Task backgroundTask = Task.CompletedTask;
        protected readonly object taskLock = new object();
        public int Count { get => count; }
        protected int count = 0;
        public bool Empty { get => Count == 0; }

        public void Queue(Action action)
        {
            lock (taskLock)
            {
                backgroundTask = backgroundTask.ContinueWith((_) => { action(); Interlocked.Decrement(ref count); });
            }
            Interlocked.Increment(ref count); 
        }
    }
}
