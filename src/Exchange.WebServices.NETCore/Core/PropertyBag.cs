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

using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Xml;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a property bag keyed on PropertyDefinition objects.
/// </summary>
internal class PropertyBag
{
    private readonly List<PropertyDefinition> _addedProperties = new();

    private readonly Dictionary<PropertyDefinition, object?> _deletedProperties = new();
    private readonly List<PropertyDefinition> _loadedProperties = new();

    private readonly List<PropertyDefinition> _modifiedProperties = new();
    private bool _isDirty;
    private bool _loading;
    private bool _onlySummaryPropertiesRequested;
    private PropertySet? _requestedPropertySet;

    /// <summary>
    ///     Gets a dictionary holding the bag's properties.
    /// </summary>
    internal Dictionary<PropertyDefinition, object> Properties { get; } = new();

    /// <summary>
    ///     Gets the owner of this bag.
    /// </summary>
    internal ServiceObject Owner { get; }

    /// <summary>
    ///     True if the bag has pending changes, false otherwise.
    /// </summary>
    internal bool IsDirty
    {
        get
        {
            var changes = _modifiedProperties.Count + _deletedProperties.Count + _addedProperties.Count;

            return changes > 0 || _isDirty;
        }
    }

    /// <summary>
    ///     Gets or sets the value of a property.
    /// </summary>
    /// <param name="propertyDefinition">The property to get or set.</param>
    /// <returns>An object representing the value of the property.</returns>
    /// <exception cref="ServiceVersionException">Raised if this property requires a later version of Exchange.</exception>
    /// <exception cref="ServiceObjectPropertyException">
    ///     Raised for get if property hasn't been assigned or loaded. Raised for
    ///     set if property cannot be updated or deleted.
    /// </exception>
    internal object this[PropertyDefinition propertyDefinition]
    {
        get
        {
            var propertyValue = GetPropertyValueOrException(propertyDefinition, out var serviceException);
            if (serviceException == null)
            {
                return propertyValue!;
            }

            throw serviceException;
        }

        set
        {
            if (propertyDefinition.Version > Owner.Service.RequestedServerVersion)
            {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.PropertyIncompatibleWithRequestVersion,
                        propertyDefinition.Name,
                        propertyDefinition.Version
                    )
                );
            }

            // If the property bag is not in the loading state, we need to verify whether
            // the property can actually be set or updated.
            if (!_loading)
            {
                // If the owner is new and if the property cannot be set, throw.
                if (Owner.IsNew &&
                    !propertyDefinition.HasFlag(PropertyDefinitionFlags.CanSet, Owner.Service.RequestedServerVersion))
                {
                    throw new ServiceObjectPropertyException(Strings.PropertyIsReadOnly, propertyDefinition);
                }

                if (!Owner.IsNew)
                {
                    // If owner is an item attachment, properties cannot be updated (EWS doesn't support updating item attachments)
                    if (Owner is Item ownerItem && ownerItem.IsAttachment)
                    {
                        throw new ServiceObjectPropertyException(
                            Strings.ItemAttachmentCannotBeUpdated,
                            propertyDefinition
                        );
                    }

                    // If the property cannot be deleted, throw.
                    if (value == null && !propertyDefinition.HasFlag(PropertyDefinitionFlags.CanDelete))
                    {
                        throw new ServiceObjectPropertyException(Strings.PropertyCannotBeDeleted, propertyDefinition);
                    }

                    // If the property cannot be updated, throw.
                    if (!propertyDefinition.HasFlag(PropertyDefinitionFlags.CanUpdate))
                    {
                        throw new ServiceObjectPropertyException(Strings.PropertyCannotBeUpdated, propertyDefinition);
                    }
                }
            }

            // If the value is set to null, delete the property.
            if (value == null)
            {
                DeleteProperty(propertyDefinition);
            }
            else
            {
                if (Properties.TryGetValue(propertyDefinition, out var currentValue))
                {
                    if (currentValue is ComplexProperty complexProperty)
                    {
                        complexProperty.OnChange -= PropertyChanged;
                    }
                }

                // If the property was to be deleted, the deletion becomes an update.
                if (_deletedProperties.Remove(propertyDefinition))
                {
                    AddToChangeList(propertyDefinition, _modifiedProperties);
                }
                else
                {
                    // If the property value was not set, we have a newly set property.
                    if (!Properties.ContainsKey(propertyDefinition))
                    {
                        AddToChangeList(propertyDefinition, _addedProperties);
                    }
                    else
                    {
                        // The last case is that we have a modified property.
                        if (!_modifiedProperties.Contains(propertyDefinition))
                        {
                            AddToChangeList(propertyDefinition, _modifiedProperties);
                        }
                    }
                }

                InitComplexProperty(value as ComplexProperty);
                Properties[propertyDefinition] = value;

                Changed();
            }
        }
    }

    /// <summary>
    ///     Initializes a new instance of PropertyBag.
    /// </summary>
    /// <param name="owner">The owner of the bag.</param>
    internal PropertyBag(ServiceObject owner)
    {
        EwsUtilities.Assert(owner != null, "PropertyBag.ctor", "owner is null");

        Owner = owner;
    }

    /// <summary>
    ///     Adds the specified property to the specified change list if it is not already present.
    /// </summary>
    /// <param name="propertyDefinition">The property to add to the change list.</param>
    /// <param name="changeList">The change list to add the property to.</param>
    internal static void AddToChangeList(PropertyDefinition propertyDefinition, List<PropertyDefinition> changeList)
    {
        if (!changeList.Contains(propertyDefinition))
        {
            changeList.Add(propertyDefinition);
        }
    }

    /// <summary>
    ///     Gets the name of the property update item.
    /// </summary>
    /// <param name="serviceObject">The service object.</param>
    /// <returns></returns>
    internal static string GetPropertyUpdateItemName(ServiceObject serviceObject)
    {
        return serviceObject is Folder ? XmlElementNames.Folder : XmlElementNames.Item;
    }

    /// <summary>
    ///     Determines whether specified property is loaded. This also includes
    ///     properties that were requested when the property bag was loaded but
    ///     were not returned by the server. In this case, the property value
    ///     will be null.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <returns>
    ///     <c>true</c> if property was loaded or requested; otherwise, <c>false</c>.
    /// </returns>
    internal bool IsPropertyLoaded(PropertyDefinition propertyDefinition)
    {
        // Is the property loaded?
        if (_loadedProperties.Contains(propertyDefinition))
        {
            return true;
        }

        // Was the property requested?
        return IsRequestedProperty(propertyDefinition);
    }

    /// <summary>
    ///     Determines whether specified property was requested.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <returns>
    ///     <c>true</c> if property was requested; otherwise, <c>false</c>.
    /// </returns>
    private bool IsRequestedProperty(PropertyDefinition propertyDefinition)
    {
        // If no requested property set, then property wasn't requested.
        if (_requestedPropertySet == null)
        {
            return false;
        }

        // If base property set is all first-class properties, use the appropriate list of
        // property definitions to see if this property was requested. Otherwise, property had 
        // to be explicitly requested and needs to be listed in AdditionalProperties.
        if (_requestedPropertySet.BasePropertySet == BasePropertySet.FirstClassProperties)
        {
            var firstClassProps = _onlySummaryPropertiesRequested ? Owner.Schema.FirstClassSummaryProperties
                : Owner.Schema.FirstClassProperties;

            return firstClassProps.Contains(propertyDefinition) || _requestedPropertySet.Contains(propertyDefinition);
        }

        return _requestedPropertySet.Contains(propertyDefinition);
    }

    /// <summary>
    ///     Determines whether the specified property has been updated.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <returns>
    ///     <c>true</c> if the specified property has been updated; otherwise, <c>false</c>.
    /// </returns>
    internal bool IsPropertyUpdated(PropertyDefinition propertyDefinition)
    {
        return _modifiedProperties.Contains(propertyDefinition) || _addedProperties.Contains(propertyDefinition);
    }

    /// <summary>
    ///     Tries to get a property value based on a property definition.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <param name="propertyValue">The property value.</param>
    /// <returns>True if property was retrieved.</returns>
    internal bool TryGetProperty(PropertyDefinition propertyDefinition, [MaybeNullWhen(false)] out object propertyValue)
    {
        propertyValue = GetPropertyValueOrException(propertyDefinition, out var serviceException);
        return serviceException == null;
    }

    /// <summary>
    ///     Tries to get a property value based on a property definition.
    /// </summary>
    /// <typeparam name="T">The types of the property.</typeparam>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <param name="propertyValue">The property value.</param>
    /// <returns>True if property was retrieved.</returns>
    internal bool TryGetProperty<T>(PropertyDefinition propertyDefinition, [MaybeNullWhen(false)] out T propertyValue)
    {
        // Verify that the type parameter and property definition's type are compatible.
        if (!typeof(T).GetTypeInfo().IsAssignableFrom(propertyDefinition.Type.GetTypeInfo()))
        {
            var errorMessage = string.Format(
                Strings.PropertyDefinitionTypeMismatch,
                EwsUtilities.GetPrintableTypeName(propertyDefinition.Type),
                EwsUtilities.GetPrintableTypeName(typeof(T))
            );
            throw new ArgumentException(errorMessage, nameof(propertyDefinition));
        }

        var result = TryGetProperty(propertyDefinition, out var value);

        propertyValue = result ? (T?)value : default;

        return result;
    }

    /// <summary>
    ///     Gets the property value.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <param name="exception">Exception that would be raised if there's an error retrieving the property.</param>
    /// <returns>Property value. May be null.</returns>
    private object? GetPropertyValueOrException(
        PropertyDefinition propertyDefinition,
        out ServiceLocalException? exception
    )
    {
        exception = null;

        if (propertyDefinition.Version > Owner.Service.RequestedServerVersion)
        {
            exception = new ServiceVersionException(
                string.Format(
                    Strings.PropertyIncompatibleWithRequestVersion,
                    propertyDefinition.Name,
                    propertyDefinition.Version
                )
            );

            return null;
        }

        if (TryGetValue(propertyDefinition, out var propertyValue))
        {
            // If the requested property is in the bag, return it.
            return propertyValue;
        }

        if (propertyDefinition.HasFlag(PropertyDefinitionFlags.AutoInstantiateOnRead))
        {
            // The requested property is an auto-instantiate-on-read property
            var complexPropertyDefinition = propertyDefinition as ComplexPropertyDefinitionBase;

            EwsUtilities.Assert(
                complexPropertyDefinition != null,
                "PropertyBag.get_this[]",
                "propertyDefinition is marked with AutoInstantiateOnRead but is not a descendant of ComplexPropertyDefinitionBase"
            );

            propertyValue = complexPropertyDefinition.CreatePropertyInstance(Owner);

            if (propertyValue != null)
            {
                InitComplexProperty(propertyValue as ComplexProperty);
                Properties[propertyDefinition] = propertyValue;
            }
        }
        else
        {
            // If the property is not the Id (we need to let developers read the Id when it's null) and if has
            // not been loaded, we throw.
            if (propertyDefinition != Owner.GetIdPropertyDefinition())
            {
                if (!IsPropertyLoaded(propertyDefinition))
                {
                    exception = new ServiceObjectPropertyException(
                        Strings.MustLoadOrAssignPropertyBeforeAccess,
                        propertyDefinition
                    );
                    return null;
                }

                // Non-nullable properties (int, bool, etc.) must be assigned or loaded; cannot return null value.
                if (!propertyDefinition.IsNullable)
                {
                    var errorMessage = IsRequestedProperty(propertyDefinition) ? Strings.ValuePropertyNotLoaded
                        : Strings.ValuePropertyNotAssigned;
                    exception = new ServiceObjectPropertyException(errorMessage, propertyDefinition);
                }
            }
        }

        return propertyValue;
    }

    /// <summary>
    ///     Sets the isDirty flag to true and triggers dispatch of the change event to the owner
    ///     of the property bag. Changed must be called whenever an operation that changes the state
    ///     of this property bag is performed (e.g. adding or removing a property).
    /// </summary>
    internal void Changed()
    {
        _isDirty = true;
        Owner.Changed();
    }

    /// <summary>
    ///     Determines whether the property bag contains a specific property.
    /// </summary>
    /// <param name="propertyDefinition">The property to check against.</param>
    /// <returns>True if the specified property is in the bag, false otherwise.</returns>
    internal bool Contains(PropertyDefinition propertyDefinition)
    {
        return Properties.ContainsKey(propertyDefinition);
    }

    /// <summary>
    ///     Tries to retrieve the value of the specified property.
    /// </summary>
    /// <param name="propertyDefinition">The property for which to retrieve a value.</param>
    /// <param name="propertyValue">If the method succeeds, contains the value of the property.</param>
    /// <returns>True if the value could be retrieved, false otherwise.</returns>
    internal bool TryGetValue(PropertyDefinition propertyDefinition, [MaybeNullWhen(false)] out object propertyValue)
    {
        return Properties.TryGetValue(propertyDefinition, out propertyValue);
    }

    /// <summary>
    ///     Handles a change event for the specified property.
    /// </summary>
    /// <param name="complexProperty">The property that changes.</param>
    internal void PropertyChanged(ComplexProperty complexProperty)
    {
        foreach (var keyValuePair in Properties)
        {
            if (keyValuePair.Value == complexProperty)
            {
                if (!_deletedProperties.ContainsKey(keyValuePair.Key))
                {
                    AddToChangeList(keyValuePair.Key, _modifiedProperties);
                    Changed();
                }
            }
        }
    }

    /// <summary>
    ///     Deletes the property from the bag.
    /// </summary>
    /// <param name="propertyDefinition">The property to delete.</param>
    internal void DeleteProperty(PropertyDefinition propertyDefinition)
    {
        if (!_deletedProperties.ContainsKey(propertyDefinition))
        {
            Properties.TryGetValue(propertyDefinition, out var propertyValue);

            Properties.Remove(propertyDefinition);
            _modifiedProperties.Remove(propertyDefinition);
            _deletedProperties.Add(propertyDefinition, propertyValue);

            if (propertyValue is ComplexProperty complexProperty)
            {
                complexProperty.OnChange -= PropertyChanged;
            }
        }
    }

    /// <summary>
    ///     Clears the bag.
    /// </summary>
    internal void Clear()
    {
        ClearChangeLog();
        Properties.Clear();
        _loadedProperties.Clear();
        _requestedPropertySet = null;
    }

    /// <summary>
    ///     Clears the bag's change log.
    /// </summary>
    internal void ClearChangeLog()
    {
        _deletedProperties.Clear();
        _modifiedProperties.Clear();
        _addedProperties.Clear();

        foreach (var keyValuePair in Properties)
        {
            if (keyValuePair.Value is ComplexProperty complexProperty)
            {
                complexProperty.ClearChangeLog();
            }
        }

        _isDirty = false;
    }

    /// <summary>
    ///     Loads properties from XML and inserts them in the bag.
    /// </summary>
    /// <param name="reader">The reader from which to read the properties.</param>
    /// <param name="clear">Indicates whether the bag should be cleared before properties are loaded.</param>
    /// <param name="requestedPropertySet">The requested property set.</param>
    /// <param name="onlySummaryPropertiesRequested">Indicates whether summary or full properties were requested.</param>
    internal void LoadFromXml(
        EwsServiceXmlReader reader,
        bool clear,
        PropertySet? requestedPropertySet,
        bool onlySummaryPropertiesRequested
    )
    {
        if (clear)
        {
            Clear();
        }

        // Put the property bag in "loading" mode. When in loading mode, no checking is done
        // when setting property values.
        _loading = true;

        _requestedPropertySet = requestedPropertySet;
        _onlySummaryPropertiesRequested = onlySummaryPropertiesRequested;

        try
        {
            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (Owner.Schema.TryGetPropertyDefinition(reader.LocalName, out var propertyDefinition))
                    {
                        propertyDefinition.LoadPropertyValueFromXml(reader, this);

                        _loadedProperties.Add(propertyDefinition);
                    }
                    else
                    {
                        reader.SkipCurrentElement();
                    }
                }
            } while (!reader.IsEndElement(XmlNamespace.Types, Owner.GetXmlElementName()));

            ClearChangeLog();
        }
        finally
        {
            _loading = false;
        }
    }

    /// <summary>
    ///     Writes the bag's properties to XML.
    /// </summary>
    /// <param name="writer">The writer to write the properties to.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteStartElement(XmlNamespace.Types, Owner.GetXmlElementName());

        foreach (var propertyDefinition in Owner.Schema)
        {
            // The following test should not be necessary since the property bag prevents
            // properties to be set if they don't have the CanSet flag, but it doesn't hurt...
            if (propertyDefinition.HasFlag(PropertyDefinitionFlags.CanSet, writer.Service.RequestedServerVersion))
            {
                if (Contains(propertyDefinition))
                {
                    propertyDefinition.WritePropertyValueToXml(writer, this, false);
                }
            }
        }

        writer.WriteEndElement();
    }

    /// <summary>
    ///     Writes the EWS update operations corresponding to the changes that occurred in the bag to XML.
    /// </summary>
    /// <param name="writer">The writer to write the updates to.</param>
    internal void WriteToXmlForUpdate(EwsServiceXmlWriter writer)
    {
        writer.WriteStartElement(XmlNamespace.Types, Owner.GetChangeXmlElementName());

        Owner.GetId().WriteToXml(writer);

        writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Updates);

        foreach (var propertyDefinition in _addedProperties)
        {
            WriteSetUpdateToXml(writer, propertyDefinition);
        }

        foreach (var propertyDefinition in _modifiedProperties)
        {
            WriteSetUpdateToXml(writer, propertyDefinition);
        }

        foreach (var property in _deletedProperties)
        {
            WriteDeleteUpdateToXml(writer, property.Key, property.Value);
        }

        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    /// <summary>
    ///     Determines whether an EWS UpdateItem/UpdateFolder call is necessary to save the changes that
    ///     occurred in the bag.
    /// </summary>
    /// <returns>True if an UpdateItem/UpdateFolder call is necessary, false otherwise.</returns>
    internal bool GetIsUpdateCallNecessary()
    {
        var propertyDefinitions = new List<PropertyDefinition>();

        propertyDefinitions.AddRange(_addedProperties);
        propertyDefinitions.AddRange(_modifiedProperties);
        propertyDefinitions.AddRange(_deletedProperties.Keys);

        foreach (var propertyDefinition in propertyDefinitions)
        {
            if (propertyDefinition.HasFlag(PropertyDefinitionFlags.CanUpdate))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    ///     Initializes a ComplexProperty instance. When a property is inserted into the bag, it needs to be
    ///     initialized in order for changes that occur on that property to be properly detected and dispatched.
    /// </summary>
    /// <param name="complexProperty">The ComplexProperty instance to initialize.</param>
    private void InitComplexProperty(ComplexProperty? complexProperty)
    {
        if (complexProperty != null)
        {
            complexProperty.OnChange += PropertyChanged;

            if (complexProperty is IOwnedProperty ownedProperty)
            {
                ownedProperty.Owner = Owner;
            }
        }
    }

    /// <summary>
    ///     Writes an EWS SetUpdate operation for the specified property.
    /// </summary>
    /// <param name="writer">The writer to write the update to.</param>
    /// <param name="propertyDefinition">The property fro which to write the update.</param>
    private void WriteSetUpdateToXml(EwsServiceXmlWriter writer, PropertyDefinition propertyDefinition)
    {
        // The following test should not be necessary since the property bag prevents
        // properties to be updated if they don't have the CanUpdate flag, but it
        // doesn't hurt...
        if (propertyDefinition.HasFlag(PropertyDefinitionFlags.CanUpdate))
        {
            var propertyValue = this[propertyDefinition];

            var handled = false;

            if (propertyValue is ICustomUpdateSerializer updateSerializer)
            {
                handled = updateSerializer.WriteSetUpdateToXml(writer, Owner, propertyDefinition);
            }

            if (!handled)
            {
                writer.WriteStartElement(XmlNamespace.Types, Owner.GetSetFieldXmlElementName());

                propertyDefinition.WriteToXml(writer);

                writer.WriteStartElement(XmlNamespace.Types, Owner.GetXmlElementName());
                propertyDefinition.WritePropertyValueToXml(writer, this, true /* isUpdateOperation */);
                writer.WriteEndElement();

                writer.WriteEndElement();
            }
        }
    }

    /// <summary>
    ///     Writes an EWS DeleteUpdate operation for the specified property.
    /// </summary>
    /// <param name="writer">The writer to write the update to.</param>
    /// <param name="propertyDefinition">The property fro which to write the update.</param>
    /// <param name="propertyValue">The current value of the property.</param>
    private void WriteDeleteUpdateToXml(
        EwsServiceXmlWriter writer,
        PropertyDefinition propertyDefinition,
        object? propertyValue
    )
    {
        // The following test should not be necessary since the property bag prevents
        // properties to be deleted (set to null) if they don't have the CanDelete flag,
        // but it doesn't hurt...
        if (propertyDefinition.HasFlag(PropertyDefinitionFlags.CanDelete))
        {
            var handled = false;

            if (propertyValue is ICustomUpdateSerializer updateSerializer)
            {
                handled = updateSerializer.WriteDeleteUpdateToXml(writer, Owner);
            }

            if (!handled)
            {
                writer.WriteStartElement(XmlNamespace.Types, Owner.GetDeleteFieldXmlElementName());
                propertyDefinition.WriteToXml(writer);
                writer.WriteEndElement();
            }
        }
    }

    /// <summary>
    ///     Validate property bag instance.
    /// </summary>
    internal void Validate()
    {
        foreach (var propertyDefinition in _addedProperties)
        {
            ValidatePropertyValue(propertyDefinition);
        }

        foreach (var propertyDefinition in _modifiedProperties)
        {
            ValidatePropertyValue(propertyDefinition);
        }
    }

    /// <summary>
    ///     Validates the property value.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    private void ValidatePropertyValue(PropertyDefinition propertyDefinition)
    {
        if (TryGetProperty(propertyDefinition, out var propertyValue))
        {
            if (propertyValue is ISelfValidate validatingValue)
            {
                validatingValue.Validate();
            }
        }
    }
}
