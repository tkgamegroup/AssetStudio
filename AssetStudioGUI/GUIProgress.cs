using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AssetStudio;

namespace AssetStudioGUI
{
    class GUIProgress : IProgress
    {
        private Action<int, int> action1;
        private Action<string> action2;

        public GUIProgress(Action<int, int> action1, Action<string> action2)
        {
            this.action1 = action1;
            this.action2 = action2;
        }
        public void Reset(string task)
        {
            action2(task);
        }

        public void Report(int current, int total)
        {
            action1(current, total);
        }
    }
}
