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

/// <summary>
///     Represents an appointment or a meeting. Properties available on appointments are defined in the AppointmentSchema
///     class.
/// </summary>
[Attachable]
[ServiceObjectDefinition(XmlElementNames.CalendarItem)]
public class Appointment : Item, ICalendarActionProvider
{
    /// <summary>
    ///     Initializes an unsaved local instance of <see cref="Appointment" />. To bind to an existing appointment, use
    ///     Appointment.Bind() instead.
    /// </summary>
    /// <param name="service">The ExchangeService instance to which this appointmtnt is bound.</param>
    public Appointment(ExchangeService service)
        : base(service)
    {
        // If we're running against Exchange 2007, we need to explicitly preset
        // the StartTimeZone property since Exchange 2007 will otherwise scope
        // start and end to UTC.
        if (service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1)
        {
            StartTimeZone = service.TimeZone;
        }
    }

    /// <summary>
    ///     Initializes a new instance of Appointment.
    /// </summary>
    /// <param name="parentAttachment">Parent attachment.</param>
    /// <param name="isNew">If true, attachment is new.</param>
    internal Appointment(ItemAttachment parentAttachment, bool isNew)
        : base(parentAttachment)
    {
        // If we're running against Exchange 2007, we need to explicitly preset
        // the StartTimeZone property since Exchange 2007 will otherwise scope
        // start and end to UTC.
        if (parentAttachment.Service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1)
        {
            if (isNew)
            {
                StartTimeZone = parentAttachment.Service.TimeZone;
            }
        }
    }

    /// <summary>
    ///     Binds to an existing appointment and loads the specified set of properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the appointment.</param>
    /// <param name="id">The Id of the appointment to bind to.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <returns>An Appointment instance representing the appointment corresponding to the specified Id.</returns>
    public static new Task<Appointment> Bind(
        ExchangeService service,
        ItemId id,
        PropertySet propertySet,
        CancellationToken token = default
    )
    {
        return service.BindToItem<Appointment>(id, propertySet, token);
    }

    /// <summary>
    ///     Binds to an existing appointment and loads its first class properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the appointment.</param>
    /// <param name="id">The Id of the appointment to bind to.</param>
    /// <returns>An Appointment instance representing the appointment corresponding to the specified Id.</returns>
    public static new Task<Appointment> Bind(ExchangeService service, ItemId id)
    {
        return Bind(service, id, PropertySet.FirstClassProperties);
    }

    /// <summary>
    ///     Binds to an occurence of an existing appointment and loads its first class properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the appointment.</param>
    /// <param name="recurringMasterId">The Id of the recurring master that the index represents an occurrence of.</param>
    /// <param name="occurenceIndex">The index of the occurrence.</param>
    /// <returns>
    ///     An Appointment instance representing the appointment occurence corresponding to the specified occurence index
    ///     .
    /// </returns>
    public static Task<Appointment> BindToOccurrence(
        ExchangeService service,
        ItemId recurringMasterId,
        int occurenceIndex
    )
    {
        return BindToOccurrence(service, recurringMasterId, occurenceIndex, PropertySet.FirstClassProperties);
    }

    /// <summary>
    ///     Binds to an occurence of an existing appointment and loads the specified set of properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the appointment.</param>
    /// <param name="recurringMasterId">The Id of the recurring master that the index represents an occurrence of.</param>
    /// <param name="occurenceIndex">The index of the occurrence.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <returns>An Appointment instance representing the appointment occurence corresponding to the specified occurence index.</returns>
    public static Task<Appointment> BindToOccurrence(
        ExchangeService service,
        ItemId recurringMasterId,
        int occurenceIndex,
        PropertySet propertySet
    )
    {
        var occurenceId = new AppointmentOccurrenceId(recurringMasterId.UniqueId, occurenceIndex);
        return Bind(service, occurenceId, propertySet);
    }

    /// <summary>
    ///     Binds to the master appointment of a recurring series and loads its first class properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the appointment.</param>
    /// <param name="occurrenceId">The Id of one of the occurrences in the series.</param>
    /// <returns>
    ///     An Appointment instance representing the master appointment of the recurring series to which the specified
    ///     occurrence belongs.
    /// </returns>
    public static Task<Appointment> BindToRecurringMaster(ExchangeService service, ItemId occurrenceId)
    {
        return BindToRecurringMaster(service, occurrenceId, PropertySet.FirstClassProperties);
    }

    /// <summary>
    ///     Binds to the master appointment of a recurring series and loads the specified set of properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the appointment.</param>
    /// <param name="occurrenceId">The Id of one of the occurrences in the series.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <returns>
    ///     An Appointment instance representing the master appointment of the recurring series to which the specified
    ///     occurrence belongs.
    /// </returns>
    public static Task<Appointment> BindToRecurringMaster(
        ExchangeService service,
        ItemId occurrenceId,
        PropertySet propertySet
    )
    {
        var recurringMasterId = new RecurringAppointmentMasterId(occurrenceId.UniqueId);
        return Bind(service, recurringMasterId, propertySet);
    }

    /// <summary>
    ///     Internal method to return the schema associated with this type of object.
    /// </summary>
    /// <returns>The schema associated with this type of object.</returns>
    internal override ServiceObjectSchema GetSchema()
    {
        return AppointmentSchema.Instance;
    }

    /// <summary>
    ///     Gets the minimum required server version.
    /// </summary>
    /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
    internal override ExchangeVersion GetMinimumRequiredServerVersion()
    {
        return ExchangeVersion.Exchange2007_SP1;
    }

    /// <summary>
    ///     Gets a value indicating whether a time zone SOAP header should be emitted in a CreateItem
    ///     or UpdateItem request so this item can be property saved or updated.
    /// </summary>
    /// <param name="isUpdateOperation">Indicates whether the operation being petrformed is an update operation.</param>
    /// <returns>
    ///     <c>true</c> if a time zone SOAP header should be emitted; otherwise, <c>false</c>.
    /// </returns>
    internal override bool GetIsTimeZoneHeaderRequired(bool isUpdateOperation)
    {
        if (isUpdateOperation)
        {
            return false;
        }

        var isStartTimeZoneSetOrUpdated = PropertyBag.IsPropertyUpdated(AppointmentSchema.StartTimeZone);
        var isEndTimeZoneSetOrUpdated = PropertyBag.IsPropertyUpdated(AppointmentSchema.EndTimeZone);

        if (isStartTimeZoneSetOrUpdated && isEndTimeZoneSetOrUpdated)
        {
            // If both StartTimeZone and EndTimeZone have been set or updated and are the same as the service's
            // time zone, we emit the time zone header and StartTimeZone and EndTimeZone are not emitted.
            TimeZoneInfo startTimeZone;
            TimeZoneInfo endTimeZone;

            PropertyBag.TryGetProperty(AppointmentSchema.StartTimeZone, out startTimeZone);
            PropertyBag.TryGetProperty(AppointmentSchema.EndTimeZone, out endTimeZone);

            return startTimeZone == Service.TimeZone || endTimeZone == Service.TimeZone;
        }

        return true;
    }

    /// <summary>
    ///     Determines whether properties defined with ScopedDateTimePropertyDefinition require custom time zone scoping.
    /// </summary>
    /// <returns>
    ///     <c>true</c> if this item type requires custom scoping for scoped date/time properties; otherwise, <c>false</c>.
    /// </returns>
    internal override bool GetIsCustomDateTimeScopingRequired()
    {
        return true;
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();

        //  Make sure that if we're on the Exchange2007_SP1 schema version, if any of the following
        //  properties are set or updated:
        //      o   Start
        //      o   End
        //      o   IsAllDayEvent
        //      o   Recurrence
        //  ... then, we must send the MeetingTimeZone element (which is generated from StartTimeZone for
        //  Exchange2007_SP1 requests (see StartTimeZonePropertyDefinition.cs).  If the StartTimeZone isn't
        //  in the property bag, then throw, because clients must supply the proper time zone - either by
        //  loading it from a currently-existing appointment, or by setting it directly.  Otherwise, to dirty
        //  the StartTimeZone property, we just set it to its current value.
        if ((Service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1) &&
            !Service.Exchange2007CompatibilityMode)
        {
            if (PropertyBag.IsPropertyUpdated(AppointmentSchema.Start) ||
                PropertyBag.IsPropertyUpdated(AppointmentSchema.End) ||
                PropertyBag.IsPropertyUpdated(AppointmentSchema.IsAllDayEvent) ||
                PropertyBag.IsPropertyUpdated(AppointmentSchema.Recurrence))
            {
                //  If the property isn't in the property bag, throw....
                if (!PropertyBag.Contains(AppointmentSchema.StartTimeZone))
                {
                    throw new ServiceLocalException(Strings.StartTimeZoneRequired);
                }

                //  Otherwise, set the time zone to its current value to force it to be sent with the request.
                StartTimeZone = StartTimeZone;
            }
        }
    }

    /// <summary>
    ///     Creates a reply response to the organizer and/or attendees of the meeting.
    /// </summary>
    /// <param name="replyAll">Indicates whether the reply should go to the organizer only or to all the attendees.</param>
    /// <returns>A ResponseMessage representing the reply response that can subsequently be modified and sent.</returns>
    public ResponseMessage CreateReply(bool replyAll)
    {
        ThrowIfThisIsNew();

        return new ResponseMessage(this, replyAll ? ResponseMessageType.ReplyAll : ResponseMessageType.Reply);
    }

    /// <summary>
    ///     Replies to the organizer and/or the attendees of the meeting. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="bodyPrefix">The prefix to prepend to the body of the meeting.</param>
    /// <param name="replyAll">Indicates whether the reply should go to the organizer only or to all the attendees.</param>
    public System.Threading.Tasks.Task Reply(MessageBody bodyPrefix, bool replyAll)
    {
        var responseMessage = CreateReply(replyAll);

        responseMessage.BodyPrefix = bodyPrefix;

        return responseMessage.SendAndSaveCopy();
    }

    /// <summary>
    ///     Creates a forward message from this appointment.
    /// </summary>
    /// <returns>A ResponseMessage representing the forward response that can subsequently be modified and sent.</returns>
    public ResponseMessage CreateForward()
    {
        ThrowIfThisIsNew();

        return new ResponseMessage(this, ResponseMessageType.Forward);
    }

    /// <summary>
    ///     Forwards the appointment. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="bodyPrefix">The prefix to prepend to the original body of the message.</param>
    /// <param name="toRecipients">The recipients to forward the appointment to.</param>
    public System.Threading.Tasks.Task Forward(MessageBody bodyPrefix, params EmailAddress[] toRecipients)
    {
        return Forward(bodyPrefix, (IEnumerable<EmailAddress>)toRecipients);
    }

    /// <summary>
    ///     Forwards the appointment. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="bodyPrefix">The prefix to prepend to the original body of the message.</param>
    /// <param name="toRecipients">The recipients to forward the appointment to.</param>
    public System.Threading.Tasks.Task Forward(MessageBody bodyPrefix, IEnumerable<EmailAddress> toRecipients)
    {
        var responseMessage = CreateForward();

        responseMessage.BodyPrefix = bodyPrefix;
        responseMessage.ToRecipients.AddRange(toRecipients);

        return responseMessage.SendAndSaveCopy();
    }

    /// <summary>
    ///     Saves this appointment in the specified folder. Calling this method results in at least one call to EWS.
    ///     Mutliple calls to EWS might be made if attachments have been added.
    /// </summary>
    /// <param name="destinationFolderName">The name of the folder in which to save this appointment.</param>
    /// <param name="sendInvitationsMode">Specifies if and how invitations should be sent if this appointment is a meeting.</param>
    public System.Threading.Tasks.Task Save(
        WellKnownFolderName destinationFolderName,
        SendInvitationsMode sendInvitationsMode,
        CancellationToken token = default
    )
    {
        return InternalCreate(new FolderId(destinationFolderName), null, sendInvitationsMode, token);
    }

    /// <summary>
    ///     Saves this appointment in the specified folder. Calling this method results in at least one call to EWS.
    ///     Mutliple calls to EWS might be made if attachments have been added.
    /// </summary>
    /// <param name="destinationFolderId">The Id of the folder in which to save this appointment.</param>
    /// <param name="sendInvitationsMode">Specifies if and how invitations should be sent if this appointment is a meeting.</param>
    public System.Threading.Tasks.Task Save(
        FolderId destinationFolderId,
        SendInvitationsMode sendInvitationsMode,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");

        return InternalCreate(destinationFolderId, null, sendInvitationsMode, token);
    }

    /// <summary>
    ///     Saves this appointment in the Calendar folder. Calling this method results in at least one call to EWS.
    ///     Mutliple calls to EWS might be made if attachments have been added.
    /// </summary>
    /// <param name="sendInvitationsMode">Specifies if and how invitations should be sent if this appointment is a meeting.</param>
    public System.Threading.Tasks.Task Save(SendInvitationsMode sendInvitationsMode, CancellationToken token = default)
    {
        return InternalCreate(null, null, sendInvitationsMode, token);
    }

    /// <summary>
    ///     Applies the local changes that have been made to this appointment. Calling this method results in at least one call
    ///     to EWS.
    ///     Mutliple calls to EWS might be made if attachments have been added or removed.
    /// </summary>
    /// <param name="conflictResolutionMode">Specifies how conflicts should be resolved.</param>
    /// <param name="sendInvitationsOrCancellationsMode">
    ///     Specifies if and how invitations or cancellations should be sent if
    ///     this appointment is a meeting.
    /// </param>
    public Task<Item?> Update(
        ConflictResolutionMode conflictResolutionMode,
        SendInvitationsOrCancellationsMode sendInvitationsOrCancellationsMode,
        CancellationToken token = default
    )
    {
        return InternalUpdate(null, conflictResolutionMode, null, sendInvitationsOrCancellationsMode, token);
    }

    /// <summary>
    ///     Deletes this appointment. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="deleteMode">The deletion mode.</param>
    /// <param name="sendCancellationsMode">Specifies if and how cancellations should be sent if this appointment is a meeting.</param>
    public Task<ServiceResponseCollection<ServiceResponse>> Delete(
        DeleteMode deleteMode,
        SendCancellationsMode sendCancellationsMode,
        CancellationToken token = default
    )
    {
        return InternalDelete(deleteMode, sendCancellationsMode, null, token);
    }

    /// <summary>
    ///     Creates a local meeting acceptance message that can be customized and sent.
    /// </summary>
    /// <param name="tentative">Specifies whether the meeting will be tentatively accepted.</param>
    /// <returns>An AcceptMeetingInvitationMessage representing the meeting acceptance message. </returns>
    public AcceptMeetingInvitationMessage CreateAcceptMessage(bool tentative)
    {
        return new AcceptMeetingInvitationMessage(this, tentative);
    }

    /// <summary>
    ///     Creates a local meeting cancellation message that can be customized and sent.
    /// </summary>
    /// <returns>A CancelMeetingMessage representing the meeting cancellation message. </returns>
    public CancelMeetingMessage CreateCancelMeetingMessage()
    {
        return new CancelMeetingMessage(this);
    }

    /// <summary>
    ///     Creates a local meeting declination message that can be customized and sent.
    /// </summary>
    /// <returns>A DeclineMeetingInvitation representing the meeting declination message. </returns>
    public DeclineMeetingInvitationMessage CreateDeclineMessage()
    {
        return new DeclineMeetingInvitationMessage(this);
    }

    /// <summary>
    ///     Accepts the meeting. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="sendResponse">Indicates whether to send a response to the organizer.</param>
    /// <returns>
    ///     A CalendarActionResults object containing the various items that were created or modified as a
    ///     results of this operation.
    /// </returns>
    public Task<CalendarActionResults> Accept(bool sendResponse)
    {
        return InternalAccept(false, sendResponse);
    }

    /// <summary>
    ///     Tentatively accepts the meeting. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="sendResponse">Indicates whether to send a response to the organizer.</param>
    /// <returns>
    ///     A CalendarActionResults object containing the various items that were created or modified as a
    ///     results of this operation.
    /// </returns>
    public Task<CalendarActionResults> AcceptTentatively(bool sendResponse)
    {
        return InternalAccept(true, sendResponse);
    }

    /// <summary>
    ///     Accepts the meeting.
    /// </summary>
    /// <param name="tentative">True if tentative accept.</param>
    /// <param name="sendResponse">Indicates whether to send a response to the organizer.</param>
    /// <returns>
    ///     A CalendarActionResults object containing the various items that were created or modified as a
    ///     results of this operation.
    /// </returns>
    internal Task<CalendarActionResults> InternalAccept(bool tentative, bool sendResponse)
    {
        var accept = CreateAcceptMessage(tentative);

        if (sendResponse)
        {
            return accept.SendAndSaveCopy();
        }

        return accept.Save();
    }

    /// <summary>
    ///     Cancels the meeting and sends cancellation messages to all attendees. Calling this method results in a call to EWS.
    /// </summary>
    /// <returns>
    ///     A CalendarActionResults object containing the various items that were created or modified as a
    ///     results of this operation.
    /// </returns>
    public Task<CalendarActionResults> CancelMeeting()
    {
        return CreateCancelMeetingMessage().SendAndSaveCopy();
    }

    /// <summary>
    ///     Cancels the meeting and sends cancellation messages to all attendees. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="cancellationMessageText">Cancellation message text sent to all attendees.</param>
    /// <returns>
    ///     A CalendarActionResults object containing the various items that were created or modified as a
    ///     results of this operation.
    /// </returns>
    public Task<CalendarActionResults> CancelMeeting(string cancellationMessageText)
    {
        var cancelMsg = CreateCancelMeetingMessage();
        cancelMsg.Body = cancellationMessageText;
        return cancelMsg.SendAndSaveCopy();
    }

    /// <summary>
    ///     Declines the meeting invitation. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="sendResponse">Indicates whether to send a response to the organizer.</param>
    /// <returns>
    ///     A CalendarActionResults object containing the various items that were created or modified as a
    ///     results of this operation.
    /// </returns>
    public Task<CalendarActionResults> Decline(bool sendResponse)
    {
        var decline = CreateDeclineMessage();

        if (sendResponse)
        {
            return decline.SendAndSaveCopy();
        }

        return decline.Save();
    }

    /// <summary>
    ///     Gets the default setting for sending cancellations on Delete.
    /// </summary>
    /// <returns>If Delete() is called on Appointment, we want to send cancellations and save a copy.</returns>
    internal override SendCancellationsMode? DefaultSendCancellationsMode => SendCancellationsMode.SendToAllAndSaveCopy;

    /// <summary>
    ///     Gets the default settings for sending invitations on Save.
    /// </summary>
    internal override SendInvitationsMode? DefaultSendInvitationsMode => SendInvitationsMode.SendToAllAndSaveCopy;

    /// <summary>
    ///     Gets the default settings for sending invitations or cancellations on Update.
    /// </summary>
    internal override SendInvitationsOrCancellationsMode? DefaultSendInvitationsOrCancellationsMode =>
        SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy;


    #region Properties

    /// <summary>
    ///     Gets or sets the start time of the appointment.
    /// </summary>
    public DateTime Start
    {
        get => (DateTime)PropertyBag[AppointmentSchema.Start];
        set => PropertyBag[AppointmentSchema.Start] = value;
    }

    /// <summary>
    ///     Gets or sets the end time of the appointment.
    /// </summary>
    public DateTime End
    {
        get => (DateTime)PropertyBag[AppointmentSchema.End];
        set => PropertyBag[AppointmentSchema.End] = value;
    }

    /// <summary>
    ///     Gets the original start time of this appointment.
    /// </summary>
    public DateTime OriginalStart => (DateTime)PropertyBag[AppointmentSchema.OriginalStart];

    /// <summary>
    ///     Gets or sets a value indicating whether this appointment is an all day event.
    /// </summary>
    public bool IsAllDayEvent
    {
        get => (bool)PropertyBag[AppointmentSchema.IsAllDayEvent];
        set => PropertyBag[AppointmentSchema.IsAllDayEvent] = value;
    }

    /// <summary>
    ///     Gets or sets a value indicating the free/busy status of the owner of this appointment.
    /// </summary>
    public LegacyFreeBusyStatus LegacyFreeBusyStatus
    {
        get => (LegacyFreeBusyStatus)PropertyBag[AppointmentSchema.LegacyFreeBusyStatus];
        set => PropertyBag[AppointmentSchema.LegacyFreeBusyStatus] = value;
    }

    /// <summary>
    ///     Gets or sets the location of this appointment.
    /// </summary>
    public string Location
    {
        get => (string)PropertyBag[AppointmentSchema.Location];
        set => PropertyBag[AppointmentSchema.Location] = value;
    }

    /// <summary>
    ///     Gets a text indicating when this appointment occurs. The text returned by When is localized using the Exchange
    ///     Server culture or using the culture specified in the PreferredCulture property of the ExchangeService object this
    ///     appointment is bound to.
    /// </summary>
    public string When => (string)PropertyBag[AppointmentSchema.When];

    /// <summary>
    ///     Gets a value indicating whether the appointment is a meeting.
    /// </summary>
    public bool IsMeeting => (bool)PropertyBag[AppointmentSchema.IsMeeting];

    /// <summary>
    ///     Gets a value indicating whether the appointment has been cancelled.
    /// </summary>
    public bool IsCancelled => (bool)PropertyBag[AppointmentSchema.IsCancelled];

    /// <summary>
    ///     Gets a value indicating whether the appointment is recurring.
    /// </summary>
    public bool IsRecurring => (bool)PropertyBag[AppointmentSchema.IsRecurring];

    /// <summary>
    ///     Gets a value indicating whether the meeting request has already been sent.
    /// </summary>
    public bool MeetingRequestWasSent => (bool)PropertyBag[AppointmentSchema.MeetingRequestWasSent];

    /// <summary>
    ///     Gets or sets a value indicating whether responses are requested when invitations are sent for this meeting.
    /// </summary>
    public bool IsResponseRequested
    {
        get => (bool)PropertyBag[AppointmentSchema.IsResponseRequested];
        set => PropertyBag[AppointmentSchema.IsResponseRequested] = value;
    }

    /// <summary>
    ///     Gets a value indicating the type of this appointment.
    /// </summary>
    public AppointmentType AppointmentType => (AppointmentType)PropertyBag[AppointmentSchema.AppointmentType];

    /// <summary>
    ///     Gets a value indicating what was the last response of the user that loaded this meeting.
    /// </summary>
    public MeetingResponseType MyResponseType => (MeetingResponseType)PropertyBag[AppointmentSchema.MyResponseType];

    /// <summary>
    ///     Gets the organizer of this meeting. The Organizer property is read-only and is only relevant for attendees.
    ///     The organizer of a meeting is automatically set to the user that created the meeting.
    /// </summary>
    public EmailAddress Organizer => (EmailAddress)PropertyBag[AppointmentSchema.Organizer];

    /// <summary>
    ///     Gets a list of required attendees for this meeting.
    /// </summary>
    public AttendeeCollection RequiredAttendees => (AttendeeCollection)PropertyBag[AppointmentSchema.RequiredAttendees];

    /// <summary>
    ///     Gets a list of optional attendeed for this meeting.
    /// </summary>
    public AttendeeCollection OptionalAttendees => (AttendeeCollection)PropertyBag[AppointmentSchema.OptionalAttendees];

    /// <summary>
    ///     Gets a list of resources for this meeting.
    /// </summary>
    public AttendeeCollection Resources => (AttendeeCollection)PropertyBag[AppointmentSchema.Resources];

    /// <summary>
    ///     Gets the number of calendar entries that conflict with this appointment in the authenticated user's calendar.
    /// </summary>
    public int ConflictingMeetingCount => (int)PropertyBag[AppointmentSchema.ConflictingMeetingCount];

    /// <summary>
    ///     Gets the number of calendar entries that are adjacent to this appointment in the authenticated user's calendar.
    /// </summary>
    public int AdjacentMeetingCount => (int)PropertyBag[AppointmentSchema.AdjacentMeetingCount];

    /// <summary>
    ///     Gets a list of meetings that conflict with this appointment in the authenticated user's calendar.
    /// </summary>
    public ItemCollection<Appointment> ConflictingMeetings =>
        (ItemCollection<Appointment>)PropertyBag[AppointmentSchema.ConflictingMeetings];

    /// <summary>
    ///     Gets a list of meetings that conflict with this appointment in the authenticated user's calendar.
    /// </summary>
    public ItemCollection<Appointment> AdjacentMeetings =>
        (ItemCollection<Appointment>)PropertyBag[AppointmentSchema.AdjacentMeetings];

    /// <summary>
    ///     Gets the duration of this appointment.
    /// </summary>
    public TimeSpan Duration => (TimeSpan)PropertyBag[AppointmentSchema.Duration];

    /// <summary>
    ///     Gets the name of the time zone this appointment is defined in.
    /// </summary>
    public string TimeZone => (string)PropertyBag[AppointmentSchema.TimeZone];

    /// <summary>
    ///     Gets the time when the attendee replied to the meeting request.
    /// </summary>
    public DateTime AppointmentReplyTime => (DateTime)PropertyBag[AppointmentSchema.AppointmentReplyTime];

    /// <summary>
    ///     Gets the sequence number of this appointment.
    /// </summary>
    public int AppointmentSequenceNumber => (int)PropertyBag[AppointmentSchema.AppointmentSequenceNumber];

    /// <summary>
    ///     Gets the state of this appointment.
    /// </summary>
    public int AppointmentState => (int)PropertyBag[AppointmentSchema.AppointmentState];

    /// <summary>
    ///     Gets or sets the recurrence pattern for this appointment. Available recurrence pattern classes include
    ///     Recurrence.DailyPattern, Recurrence.MonthlyPattern and Recurrence.YearlyPattern.
    /// </summary>
    public Recurrence Recurrence
    {
        get => (Recurrence)PropertyBag[AppointmentSchema.Recurrence];

        set
        {
            if (value != null)
            {
                if (value.IsRegenerationPattern)
                {
                    throw new ServiceLocalException(Strings.RegenerationPatternsOnlyValidForTasks);
                }
            }

            PropertyBag[AppointmentSchema.Recurrence] = value;
        }
    }

    /// <summary>
    ///     Gets an OccurrenceInfo identifying the first occurrence of this meeting.
    /// </summary>
    public OccurrenceInfo FirstOccurrence => (OccurrenceInfo)PropertyBag[AppointmentSchema.FirstOccurrence];

    /// <summary>
    ///     Gets an OccurrenceInfo identifying the last occurrence of this meeting.
    /// </summary>
    public OccurrenceInfo LastOccurrence => (OccurrenceInfo)PropertyBag[AppointmentSchema.LastOccurrence];

    /// <summary>
    ///     Gets a list of modified occurrences for this meeting.
    /// </summary>
    public OccurrenceInfoCollection ModifiedOccurrences =>
        (OccurrenceInfoCollection)PropertyBag[AppointmentSchema.ModifiedOccurrences];

    /// <summary>
    ///     Gets a list of deleted occurrences for this meeting.
    /// </summary>
    public DeletedOccurrenceInfoCollection DeletedOccurrences =>
        (DeletedOccurrenceInfoCollection)PropertyBag[AppointmentSchema.DeletedOccurrences];

    /// <summary>
    ///     Gets or sets time zone of the start property of this appointment.
    /// </summary>
    public TimeZoneInfo StartTimeZone
    {
        get => (TimeZoneInfo)PropertyBag[AppointmentSchema.StartTimeZone];
        set => PropertyBag[AppointmentSchema.StartTimeZone] = value;
    }

    /// <summary>
    ///     Gets or sets time zone of the end property of this appointment.
    /// </summary>
    public TimeZoneInfo EndTimeZone
    {
        get => (TimeZoneInfo)PropertyBag[AppointmentSchema.EndTimeZone];
        set => PropertyBag[AppointmentSchema.EndTimeZone] = value;
    }

    /// <summary>
    ///     Gets or sets the type of conferencing that will be used during the meeting.
    /// </summary>
    public int ConferenceType
    {
        get => (int)PropertyBag[AppointmentSchema.ConferenceType];
        set => PropertyBag[AppointmentSchema.ConferenceType] = value;
    }

    /// <summary>
    ///     Gets or sets a value indicating whether new time proposals are allowed for attendees of this meeting.
    /// </summary>
    public bool AllowNewTimeProposal
    {
        get => (bool)PropertyBag[AppointmentSchema.AllowNewTimeProposal];
        set => PropertyBag[AppointmentSchema.AllowNewTimeProposal] = value;
    }

    /// <summary>
    ///     Gets or sets a value indicating whether this is an online meeting.
    /// </summary>
    public bool IsOnlineMeeting
    {
        get => (bool)PropertyBag[AppointmentSchema.IsOnlineMeeting];
        set => PropertyBag[AppointmentSchema.IsOnlineMeeting] = value;
    }

    /// <summary>
    ///     Gets or sets the URL of the meeting workspace. A meeting workspace is a shared Web site for planning meetings and
    ///     tracking results.
    /// </summary>
    public string MeetingWorkspaceUrl
    {
        get => (string)PropertyBag[AppointmentSchema.MeetingWorkspaceUrl];
        set => PropertyBag[AppointmentSchema.MeetingWorkspaceUrl] = value;
    }

    /// <summary>
    ///     Gets or sets the URL of the Microsoft NetShow online meeting.
    /// </summary>
    public string NetShowUrl
    {
        get => (string)PropertyBag[AppointmentSchema.NetShowUrl];
        set => PropertyBag[AppointmentSchema.NetShowUrl] = value;
    }

    /// <summary>
    ///     Gets or sets the ICalendar Uid.
    /// </summary>
    public string ICalUid
    {
        get => (string)PropertyBag[AppointmentSchema.ICalUid];
        set => PropertyBag[AppointmentSchema.ICalUid] = value;
    }

    /// <summary>
    ///     Gets the ICalendar RecurrenceId.
    /// </summary>
    public DateTime? ICalRecurrenceId => (DateTime?)PropertyBag[AppointmentSchema.ICalRecurrenceId];

    /// <summary>
    ///     Gets the ICalendar DateTimeStamp.
    /// </summary>
    public DateTime? ICalDateTimeStamp => (DateTime?)PropertyBag[AppointmentSchema.ICalDateTimeStamp];

    /// <summary>
    ///     Gets or sets the Enhanced location object.
    /// </summary>
    public EnhancedLocation EnhancedLocation
    {
        get => (EnhancedLocation)PropertyBag[AppointmentSchema.EnhancedLocation];
        set => PropertyBag[AppointmentSchema.EnhancedLocation] = value;
    }

    /// <summary>
    ///     Gets the Url for joining an online meeting
    /// </summary>
    public string JoinOnlineMeetingUrl => (string)PropertyBag[AppointmentSchema.JoinOnlineMeetingUrl];

    /// <summary>
    ///     Gets the Online Meeting Settings
    /// </summary>
    public OnlineMeetingSettings OnlineMeetingSettings =>
        (OnlineMeetingSettings)PropertyBag[AppointmentSchema.OnlineMeetingSettings];

    #endregion
}
