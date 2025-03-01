using System;

namespace Scale_Program.Functions
{
    public class BasculaEventArgs : EventArgs
    {
        public bool IsStable { get; set; }

        public double Value { get; set; }

        public override string ToString()
        {
            return Value.ToString();
        }
    }
}