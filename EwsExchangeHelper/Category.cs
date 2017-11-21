using System;
using System.Linq;
using EwsExchangeHelper.ExchangeItems;

namespace EwsExchangeHelper
{
    public partial class MsExchangeServices
    {
        /// <summary>
        /// Getting all available Categories
        /// </summary>
        /// <returns> List of Categories</returns>
        public MasterCategoryList GetCategoryList()
        {
            return MasterCategoryList.Bind(ExchangeService);
        }

        /// <summary>
        /// Checking, if the specified categories exists
        /// </summary>
        /// <param name="categoryName">Category name</param>
        /// <param name="list">list of categories</param>
        /// <returns>true if exists, otherwise false</returns>
        public bool ExistsCategory(string categoryName, MasterCategoryList list = null)
        {
            if (list == null)
                list = GetCategoryList();

            return list.Categories.Any(category => category.Name == categoryName);
        }


        /// <summary>
        /// Check if the categories exists, otherwise create it
        /// </summary>
        /// <param name="categoryName">name of the category</param>
        public void CheckAndCreateCategory(string categoryName)
        {
            MasterCategoryList list;

            try
            {
                list = GetCategoryList();
            }
            catch
            {
                list = MasterCategoryList.BindOrCreate(ExchangeService);
            }

            if (ExistsCategory(categoryName, list)) return;

            list.Categories.Add(new Category
            {
                Name = categoryName,
                Id = Guid.NewGuid()
            });
            list.Update();
        }
    }
}