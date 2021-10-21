using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AssetStudio
{
    public interface IProgress
    {
        void Reset(string task);
        void Report(int current, int total);
    }

    public sealed class DummyProgress : IProgress
    {
        public void Reset(string task) { }
        public void Report(int current, int total) { }
    }
}
