namespace FilterExcel.Infrastructure.Helpers;

public class RegexOperationsHelper : IRegexOperationsHelper
{
    public int CalculateExpression(string expression)
    {
        var elements = Regex.Split(expression, @"([-+])").Select(e => e.Trim()).ToList();
        var result = int.Parse(elements[0]);

        for (int i = 1; i < elements.Count; i += 2)
        {
            var op = elements[i];
            var num = int.Parse(elements[i + 1]);

            result = op == "+" ? result + num : result - num;
        }

        return result;
    }
    
    public int SumSubtractedNumbers(string expression)
    {
        var elements = Regex.Split(expression, @"([-+])").Select(e => e.Trim()).ToList();
        var sum = 0;

        for (var i = 1; i < elements.Count; i += 2)
        {
            if (elements[i] == "-")
            {
                sum += Math.Abs(int.Parse(elements[i + 1]));
            }
        }

        return sum;
    }
}