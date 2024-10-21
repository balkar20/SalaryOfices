

namespace SalaryReport
{
    public class SetCurencyInput
    {
        public  string Input { get; set; }
        public  string Currency { get; set; }

        public SetCurencyInput(string input, string currency)
        {
            Input = input;
            Currency = currency;
        }
    }
}