// <copyright file="WorkBookService.cs" company="Nexcel Solutions Vietnam">
// Copyright (c) Nexcel Solutions Vietnam. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------

namespace Fanex.Workbook.Service
{
    using System;
    using System.Collections.Generic;
    using Fanex.Workbook.Service.Business;
    using Fanex.Workbook.Service.Entities;
    using Fanex.Workbook.Service.Extensions;

    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in both code and config file together.
    /// <summary>
    /// Work book service.
    /// </summary>
    public class WorkBookService : IWorkBookService
    {
        public string GetData(int value)
        {
            LogExtensions.Log("what the hell");
            return string.Format("You entered: {0}", value);
        }

        public byte[] CreateDocument(List<DocumentSetting> input)
        {
            if (input == null) 
            {
                return null;
            } 
            
            return WorkBookManager.CreateDocument(input);
        }

        /// <summary>
        /// Get composite.
        /// </summary>
        /// <param name="composite"></param>
        /// <returns>Composite Type.</returns>
        public CompositeType GetDataUsingDataContract(CompositeType composite)
        {
            if (composite == null)
            {
                throw new ArgumentNullException("composite");
            }

            if (composite.BoolValue)
            {
                composite.StringValue += "Suffix";
            }

            return composite;
        }
    }
}
