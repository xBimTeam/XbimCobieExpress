using DocumentFormat.OpenXml;
using System;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;

namespace Xbim.IO.Table
{
    public static class WorkbookExtensions
    {
        /// <summary>
        /// Gets a <typeparamref name="T"/> using the <paramref name="getterPredicate"/>, creating a new instance with 
        /// <paramref name="createFn"/> where no item matches
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="workbook"></param>
        /// <param name="createFn"></param>
        /// <param name="getterPredicate"></param>
        /// <returns></returns>
        public static T GetOrCreate<T>(this OpenXmlElement workbook, Func<OpenXmlElement, T> createFn, Func<OpenXmlElement, T> getterPredicate = null) where T : OpenXmlElement
        {
            getterPredicate ??= (wb => wb.GetFirstChild<T>());

            T item = getterPredicate(workbook);
            item ??= createFn(workbook);
            return item;
        }

        public static TOut GetOrCreatePart<TIn, TOut>(this TIn container, Func<TIn, TOut> createFn, Func<TIn, TOut> getterPredicate = null) 
            where TIn: OpenXmlPartContainer 
            where TOut : OpenXmlPart
        {
            getterPredicate ??= (wb => wb.GetPartsOfType<TOut>().FirstOrDefault());

            TOut item = getterPredicate(container);
            item ??= createFn(container);
            return item;
        }
    }
}
