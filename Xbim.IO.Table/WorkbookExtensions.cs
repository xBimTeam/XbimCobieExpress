using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
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
            where TIn : OpenXmlCompositeElement
            where TOut : OpenXmlElement
        {
            getterPredicate ??= (wb => wb.GetFirstChild<TOut>());

            TOut item = getterPredicate(parentElement);
            item ??= createFn(parentElement);
            return item;
        }

        public static TOut GetOrCreatePart<TIn, TOut>(this TIn container, Func<TIn, TOut> createFn, Func<TIn, TOut> getterPredicate = null)
            where TIn : OpenXmlPartContainer
            where TOut : OpenXmlPart
        {
            getterPredicate ??= (wb => wb.GetPartsOfType<TOut>().FirstOrDefault());

            TOut item = getterPredicate(container);
            item ??= createFn(container);
            return item;
        }

        /// <summary>
        /// Gets the first child element of the required type, inserting a new one at the appropriate index when none
        /// exists
        /// </summary>
        /// <remarks>The order of Worksheet children is important. Incorrect ordering leads to corrupt Excel files</remarks>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        public static T GetOrCreateWorksheetChildCollection<T>(this Worksheet worksheet) where T : OpenXmlCompositeElement, new()
        {
            T collection = worksheet.GetFirstChild<T>();
            if (collection == null)
            {
                collection = new T();
                if (!worksheet.HasChildren)
                {
                    worksheet.AppendChild(collection);
                }
                else
                {
                    int collectionSchemaPos = getChildElementOrderIndex(collection);
                    int insertPos = 0;
                    int lastOrderNum = -1;
                    for (int i = 0; i < worksheet.ChildElements.Count; ++i)
                    {
                        int thisOrderNum = getChildElementOrderIndex(worksheet.ChildElements[i]);
                        if (thisOrderNum <= lastOrderNum)
                            throw new InvalidOperationException($"Internal: worksheet parts {_childElementNamesSequence[lastOrderNum]} and {_childElementNamesSequence[thisOrderNum]} out of order");
                        lastOrderNum = thisOrderNum;
                        if (thisOrderNum < collectionSchemaPos)
                            ++insertPos;
                    }
                    // this is the index to insert the new element
                    worksheet.InsertAt(collection, insertPos);
                }
            }
            return collection;
        }

        private static int getChildElementOrderIndex(OpenXmlElement collection)
        {
            int orderIndex = _childElementNamesSequence.IndexOf(collection.LocalName);
            if (orderIndex < 0)
                throw new InvalidOperationException($"Internal: worksheet part {collection.LocalName} not found");
            return orderIndex;
        }


        private static readonly List<string> _childElementNamesSequence = new List<string>()
        {
            "sheetPr",
            "dimension",
            "sheetViews",
            "sheetFormatPr",
            "cols",
            "sheetData",
            "sheetCalcPr",
            "sheetProtection",
            "protectedRanges",
            "scenarios",
            "autoFilter",
            "sortState",
            "dataConsolidate",
            "customSheetViews",
            "mergeCells",
            "phoneticPr",
            "conditionalFormatting",
            "dataValidations",
            "hyperlinks",
            "printOptions",
            "pageMargins",
            "pageSetup",
            "headerFooter",
            "rowBreaks",
            "colBreaks",
            "customProperties",
            "cellWatches",
            "ignoredErrors",
            "smartTags",
            "drawing",
            "drawingHF",
            "picture",
            "oleObjects",
            "controls",
            "webPublishItems",
            "tableParts",
            "extLst"

        };
    }
}
