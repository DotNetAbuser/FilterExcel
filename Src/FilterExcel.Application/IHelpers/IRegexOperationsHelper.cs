namespace FilterExcel.Application.IHelpers;

public interface IRegexOperationsHelper
{
    int CalculateExpression(string expression);
    int SumSubtractedNumbers(string expression);
}