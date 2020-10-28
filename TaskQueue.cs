using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShortcutSync
{
    class TaskQueue
    {
        protected Task backgroundTask = Task.CompletedTask;
        public int Count { get; protected set; } = 0;
        public bool Empty { get => Count == 0; }

        public void Queue(Action action)
        {
            backgroundTask = backgroundTask.ContinueWith((_) => { action(); --Count; });
            ++Count;
        }
    }
}
