using System;

namespace ShortcutSync
{
    public abstract class Shortcut<T>
    {
        internal Shortcut() { }

        public T TargetObject { get; protected set; }
        public string ShortcutPath { get; protected set; }

        public abstract DateTime LastUpdated { get; protected set; }
        public abstract bool Exists { get; }
        public abstract string Name { get; set; }
        public abstract bool IsValid { get; }
        public abstract bool IsOutdated { get; }

        public abstract bool Move(params string[] paths);
        public abstract bool Create();
        public abstract bool Remove();
        public abstract bool Update(bool forceUpdate = false);
        public abstract bool Update(T targetObject);
        public void CreateOrUpdate()
        {
            if (Exists)
                Update();
            else if (IsValid)
                Create();
        }
    }
}
