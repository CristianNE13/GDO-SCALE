using System.Diagnostics;
using System.Threading;
using log4net;

namespace Scale_Program.Functions
{
    public sealed class SeaLevelController
    {
        private const int SeaLevelScanInterval = 100;
        private const int MinSeaLevelScanInterval = 75;
        private readonly ILog Logger = LogManager.GetLogger("SeaLevelController");

        private Timer _timer;
        private Stopwatch stopWatch;


        public bool EnableRaisingEvents
        {
            get => _timer != null;
            set
            {
                if (value)
                {
                    _timer = new Timer(CheckForIOChanges, null, Timeout.Infinite, Timeout.Infinite);
                    _timer.Change(SeaLevelScanInterval, Timeout.Infinite);
                }
                else
                {
                    if (_timer != null)
                    {
                        _timer.Change(Timeout.Infinite, Timeout.Infinite);
                        _timer.Dispose();
                        _timer = null;
                    }
                }
            }
        }

        private void CheckForIOChanges(object stateInfo)
        {
            stopWatch.Restart();


            stopWatch.Stop();
            if (stopWatch.ElapsedMilliseconds < SeaLevelScanInterval)
                _timer.Change(SeaLevelScanInterval - stopWatch.ElapsedMilliseconds, Timeout.Infinite);
            else
                _timer.Change(MinSeaLevelScanInterval, Timeout.Infinite);
        }
    }
}