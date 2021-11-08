namespace AssetStudio
{
    public static class Progress
    {
        public static IProgress Default = new DummyProgress();

        public static void Reset(string task)
        {
            Default.Reset(task);
        }

        public static void Report(int current, int total)
        {
            Default.Report(current, total);
        }
    }
}
