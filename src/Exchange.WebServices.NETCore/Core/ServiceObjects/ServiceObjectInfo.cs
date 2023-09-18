/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data;

internal delegate object CreateServiceObjectWithServiceParam(ExchangeService srv);

internal delegate object CreateServiceObjectWithAttachmentParam(ItemAttachment itemAttachment, bool isNew);

/// <summary>
///     ServiceObjectInfo contains metadata on how to map from an element name to a ServiceObject type
///     as well as how to map from a ServiceObject type to appropriate constructors.
/// </summary>
internal class ServiceObjectInfo
{
    /// <summary>
    ///     Default constructor
    /// </summary>
    internal ServiceObjectInfo()
    {
        InitializeServiceObjectClassMap();
    }

    /// <summary>
    ///     Initializes the service object class map.
    /// </summary>
    /// <remarks>
    ///     If you add a new ServiceObject subclass that can be returned by the Server, add the type
    ///     to the class map as well as associated delegate(s) to call the constructor(s).
    /// </remarks>
    private void InitializeServiceObjectClassMap()
    {
        // Appointment
        AddServiceObjectType(
            XmlElementNames.CalendarItem,
            typeof(Appointment),
            srv => new Appointment(srv),
            (itemAttachment, isNew) => new Appointment(itemAttachment, isNew)
        );

        // CalendarFolder
        AddServiceObjectType(
            XmlElementNames.CalendarFolder,
            typeof(CalendarFolder),
            srv => new CalendarFolder(srv),
            null
        );

        // Contact
        AddServiceObjectType(
            XmlElementNames.Contact,
            typeof(Contact),
            srv => new Contact(srv),
            (itemAttachment, _) => new Contact(itemAttachment)
        );

        // ContactsFolder
        AddServiceObjectType(
            XmlElementNames.ContactsFolder,
            typeof(ContactsFolder),
            srv => new ContactsFolder(srv),
            null
        );

        // ContactGroup
        AddServiceObjectType(
            XmlElementNames.DistributionList,
            typeof(ContactGroup),
            srv => new ContactGroup(srv),
            (itemAttachment, _) => new ContactGroup(itemAttachment)
        );

        // Conversation
        AddServiceObjectType(XmlElementNames.Conversation, typeof(Conversation), srv => new Conversation(srv), null);

        // EmailMessage
        AddServiceObjectType(
            XmlElementNames.Message,
            typeof(EmailMessage),
            srv => new EmailMessage(srv),
            (itemAttachment, _) => new EmailMessage(itemAttachment)
        );

        // Folder
        AddServiceObjectType(XmlElementNames.Folder, typeof(Folder), srv => new Folder(srv), null);

        // Item
        AddServiceObjectType(
            XmlElementNames.Item,
            typeof(Item),
            srv => new Item(srv),
            (itemAttachment, _) => new Item(itemAttachment)
        );

        // MeetingCancellation
        AddServiceObjectType(
            XmlElementNames.MeetingCancellation,
            typeof(MeetingCancellation),
            srv => new MeetingCancellation(srv),
            (itemAttachment, _) => new MeetingCancellation(itemAttachment)
        );

        // MeetingMessage
        AddServiceObjectType(
            XmlElementNames.MeetingMessage,
            typeof(MeetingMessage),
            srv => new MeetingMessage(srv),
            (itemAttachment, _) => new MeetingMessage(itemAttachment)
        );

        // MeetingRequest
        AddServiceObjectType(
            XmlElementNames.MeetingRequest,
            typeof(MeetingRequest),
            srv => new MeetingRequest(srv),
            (itemAttachment, _) => new MeetingRequest(itemAttachment)
        );

        // MeetingResponse
        AddServiceObjectType(
            XmlElementNames.MeetingResponse,
            typeof(MeetingResponse),
            srv => new MeetingResponse(srv),
            (itemAttachment, _) => new MeetingResponse(itemAttachment)
        );

        // Persona
        AddServiceObjectType(XmlElementNames.Persona, typeof(Persona), srv => new Persona(srv), null);

        // PostItem
        AddServiceObjectType(
            XmlElementNames.PostItem,
            typeof(PostItem),
            srv => new PostItem(srv),
            (itemAttachment, _) => new PostItem(itemAttachment)
        );

        // SearchFolder
        AddServiceObjectType(XmlElementNames.SearchFolder, typeof(SearchFolder), srv => new SearchFolder(srv), null);

        // Task
        AddServiceObjectType(
            XmlElementNames.Task,
            typeof(Task),
            srv => new Task(srv),
            (itemAttachment, _) => new Task(itemAttachment)
        );

        // TasksFolder
        AddServiceObjectType(XmlElementNames.TasksFolder, typeof(TasksFolder), srv => new TasksFolder(srv), null);
    }

    /// <summary>
    ///     Adds specified type of service object to map.
    /// </summary>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <param name="type">The ServiceObject type.</param>
    /// <param name="createServiceObjectWithServiceParam">Delegate to create service object with service param.</param>
    /// <param name="createServiceObjectWithAttachmentParam">Delegate to create service object with attachment param.</param>
    private void AddServiceObjectType(
        string xmlElementName,
        Type type,
        CreateServiceObjectWithServiceParam createServiceObjectWithServiceParam,
        CreateServiceObjectWithAttachmentParam? createServiceObjectWithAttachmentParam
    )
    {
        XmlElementNameToServiceObjectClassMap.Add(xmlElementName, type);
        ServiceObjectConstructorsWithServiceParam.Add(type, createServiceObjectWithServiceParam);

        if (createServiceObjectWithAttachmentParam != null)
        {
            ServiceObjectConstructorsWithAttachmentParam.Add(type, createServiceObjectWithAttachmentParam);
        }
    }

    /// <summary>
    ///     Return Dictionary that maps from element name to ServiceObject Type.
    /// </summary>
    internal Dictionary<string, Type> XmlElementNameToServiceObjectClassMap { get; } = new();

    /// <summary>
    ///     Return Dictionary that maps from ServiceObject Type to CreateServiceObjectWithServiceParam delegate with
    ///     ExchangeService parameter.
    /// </summary>
    internal Dictionary<Type, CreateServiceObjectWithServiceParam> ServiceObjectConstructorsWithServiceParam { get; } =
        new();

    /// <summary>
    ///     Return Dictionary that maps from ServiceObject Type to CreateServiceObjectWithAttachmentParam delegate with
    ///     ItemAttachment parameter.
    /// </summary>
    internal Dictionary<Type, CreateServiceObjectWithAttachmentParam> ServiceObjectConstructorsWithAttachmentParam
    {
        get;
    } = new();
}
