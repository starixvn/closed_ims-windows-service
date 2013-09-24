// --------------------------------------------------------------------------
// <copyright file="IWorkbookService.cs" company="Nexcel Solutions Vietnam">
//     Copyright (c) Nexcel Solutions Vietnam. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------

namespace Fanex.Workbook.Service
{
    using System.Runtime.Serialization;
    using System.ServiceModel;
    using Fanex.Workbook.Service.Entities;

    /// <summary>
    /// Interface IWorkBookService.
    /// </summary>
    [ServiceContract]
    public interface IWorkBookService
    {
        /// <summary>
        /// Gets the data.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>System > String.</returns>
        [OperationContract]
        string GetData(int value);

        /// <summary>
        /// Creates the document.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <returns>System > Byte[][].</returns>
        [OperationContract]
        byte[] CreateDocument(DocumentFile input);

        /// <summary>
        /// Gets the data using data contract.
        /// </summary>
        /// <param name="composite">The composite.</param>
        /// <returns>Return CompositeType.</returns>
        [OperationContract]
        CompositeType GetDataUsingDataContract(CompositeType composite);

        // TODO: Add your service operations here
    }

    /// <summary>
    /// Use a data contract as illustrated in the sample below to add composite types to service operations.
    /// </summary>
    [DataContract]
    public class CompositeType
    {
        private bool _boolValue = true;
        private string _stringValue = "Hello ";

        /// <summary>
        /// Gets or sets a value indicating whether [BOOL value].
        /// </summary>
        /// <value><c>true</c> if [BOOL value]; otherwise, <c>false</c>.</value>
        [DataMember]
        public bool BoolValue
        {
            get { return _boolValue; }
            set { _boolValue = value; }
        }

        /// <summary>
        /// Gets or sets the string value.
        /// </summary>
        /// <value>The string value.</value>
        [DataMember]
        public string StringValue
        {
            get { return _stringValue; }
            set { _stringValue = value; }
        }
    }
}
