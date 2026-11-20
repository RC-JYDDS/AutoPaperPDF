using System;
using System.IO;
using System.Reflection;
using System.Collections.Generic;

namespace AutoPaper
{
    public static class EmbeddedDependencies
    {
        static bool registered;
        static readonly Dictionary<string,Assembly> cache = new Dictionary<string,Assembly>(StringComparer.OrdinalIgnoreCase);
        public static void Register()
        {
            if (registered) return;
            registered = true;
            AppDomain.CurrentDomain.AssemblyResolve += delegate(object sender, ResolveEventArgs args)
            {
                string name = new AssemblyName(args.Name).Name;
                if (String.IsNullOrEmpty(name) || name.IndexOfAny(new[]{'/', '\\', ':'}) >= 0) return null;
                lock(cache)
                {
                Assembly loaded;
                if(cache.TryGetValue(name,out loaded)) return loaded;
                var owner = typeof(EmbeddedDependencies).Assembly;
                using (Stream stream = owner.GetManifestResourceStream("AutoPaper.Dependencies." + name + ".dll"))
                {
                    if (stream == null) return null;
                    using (var memory = new MemoryStream())
                    { stream.CopyTo(memory); loaded=Assembly.Load(memory.ToArray()); cache[name]=loaded; return loaded; }
                }
                }
            };
        }
    }
}
