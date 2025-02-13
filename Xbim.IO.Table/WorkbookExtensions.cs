using DocumentFormat.OpenXml;
using System;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;

namespace Xbim.IO.Table
{
    public static class WorkbookExtensions
    {
        /// <summary>
        /// Gets a <typeparamref name="TOut"/> using the <paramref name="getterPredicate"/>, creating a new instance with 
        /// <paramref name="createFn"/> where no item matches
        /// </summary>
        /// <typeparam name="TIn"></typeparam>
        /// <typeparam name="TOut"></typeparam>
        /// <param name="parentElement"></param>
        /// <param name="createFn"></param>
        /// <param name="getterPredicate"></param>
        /// <returns></returns>
        public static TOut GetOrCreate<TIn, TOut>(this TIn parentElement, Func<TIn, TOut> createFn, Func<TIn, TOut> getterPredicate = null)
            where TIn: OpenXmlCompositeElement
            where TOut : OpenXmlElement
        {
            getterPredicate ??= (wb => wb.GetFirstChild<TOut>());

            TOut item = getterPredicate(parentElement);
            item ??= createFn(parentElement);
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
