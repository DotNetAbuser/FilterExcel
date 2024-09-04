namespace FilterExcel.Infrastructure.Extensions;

public static class ServicesCollectionExtensions
{
    public static void AddServices(this IServiceCollection services)
    {
        services
            .AddTransient<IFilterService, FilterService>()
            .AddTransient<ISearchValuesFromTxtService, SearchValuesFromTxtService>()
            .AddTransient<ISearchBySupplierService, SearchBySupplierService>()
            .AddTransient<ICountsService, CountsService>();
    }
    
    public static void AddHelpers(this IServiceCollection services)
    {
        services
            .AddTransient<IFileHelper, FileHelper>()
            .AddTransient<IExcelOperationsHelper, ExcelOperationsHelper>()
            .AddTransient<IRegexOperationsHelper, RegexOperationsHelper>();
    }
}