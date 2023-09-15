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

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a permission on a folder.
/// </summary>
[PublicAPI]
public sealed class FolderPermission : ComplexProperty
{
    #region Default permissions

    private static readonly LazyMember<Dictionary<FolderPermissionLevel, FolderPermission>> DefaultPermissions = new(
        () => new Dictionary<FolderPermissionLevel, FolderPermission>
        {
            {
                FolderPermissionLevel.None, new FolderPermission
                {
                    _canCreateItems = false,
                    _canCreateSubFolders = false,
                    _deleteItems = PermissionScope.None,
                    _editItems = PermissionScope.None,
                    _isFolderContact = false,
                    _isFolderOwner = false,
                    _isFolderVisible = false,
                    _readItems = FolderPermissionReadAccess.None,
                }
            },
            {
                FolderPermissionLevel.Contributor, new FolderPermission
                {
                    _canCreateItems = true,
                    _canCreateSubFolders = false,
                    _deleteItems = PermissionScope.None,
                    _editItems = PermissionScope.None,
                    _isFolderContact = false,
                    _isFolderOwner = false,
                    _isFolderVisible = true,
                    _readItems = FolderPermissionReadAccess.None,
                }
            },
            {
                FolderPermissionLevel.Reviewer, new FolderPermission
                {
                    _canCreateItems = false,
                    _canCreateSubFolders = false,
                    _deleteItems = PermissionScope.None,
                    _editItems = PermissionScope.None,
                    _isFolderContact = false,
                    _isFolderOwner = false,
                    _isFolderVisible = true,
                    _readItems = FolderPermissionReadAccess.FullDetails,
                }
            },
            {
                FolderPermissionLevel.NoneditingAuthor, new FolderPermission
                {
                    _canCreateItems = true,
                    _canCreateSubFolders = false,
                    _deleteItems = PermissionScope.Owned,
                    _editItems = PermissionScope.None,
                    _isFolderContact = false,
                    _isFolderOwner = false,
                    _isFolderVisible = true,
                    _readItems = FolderPermissionReadAccess.FullDetails,
                }
            },
            {
                FolderPermissionLevel.Author, new FolderPermission
                {
                    _canCreateItems = true,
                    _canCreateSubFolders = false,
                    _deleteItems = PermissionScope.Owned,
                    _editItems = PermissionScope.Owned,
                    _isFolderContact = false,
                    _isFolderOwner = false,
                    _isFolderVisible = true,
                    _readItems = FolderPermissionReadAccess.FullDetails,
                }
            },
            {
                FolderPermissionLevel.PublishingAuthor, new FolderPermission
                {
                    _canCreateItems = true,
                    _canCreateSubFolders = true,
                    _deleteItems = PermissionScope.Owned,
                    _editItems = PermissionScope.Owned,
                    _isFolderContact = false,
                    _isFolderOwner = false,
                    _isFolderVisible = true,
                    _readItems = FolderPermissionReadAccess.FullDetails,
                }
            },
            {
                FolderPermissionLevel.Editor, new FolderPermission
                {
                    _canCreateItems = true,
                    _canCreateSubFolders = false,
                    _deleteItems = PermissionScope.All,
                    _editItems = PermissionScope.All,
                    _isFolderContact = false,
                    _isFolderOwner = false,
                    _isFolderVisible = true,
                    _readItems = FolderPermissionReadAccess.FullDetails,
                }
            },
            {
                FolderPermissionLevel.PublishingEditor, new FolderPermission
                {
                    _canCreateItems = true,
                    _canCreateSubFolders = true,
                    _deleteItems = PermissionScope.All,
                    _editItems = PermissionScope.All,
                    _isFolderContact = false,
                    _isFolderOwner = false,
                    _isFolderVisible = true,
                    _readItems = FolderPermissionReadAccess.FullDetails,
                }
            },
            {
                FolderPermissionLevel.Owner, new FolderPermission
                {
                    _canCreateItems = true,
                    _canCreateSubFolders = true,
                    _deleteItems = PermissionScope.All,
                    _editItems = PermissionScope.All,
                    _isFolderContact = true,
                    _isFolderOwner = true,
                    _isFolderVisible = true,
                    _readItems = FolderPermissionReadAccess.FullDetails,
                }
            },
            {
                FolderPermissionLevel.FreeBusyTimeOnly, new FolderPermission
                {
                    _canCreateItems = false,
                    _canCreateSubFolders = false,
                    _deleteItems = PermissionScope.None,
                    _editItems = PermissionScope.None,
                    _isFolderContact = false,
                    _isFolderOwner = false,
                    _isFolderVisible = false,
                    _readItems = FolderPermissionReadAccess.TimeOnly,
                }
            },
            {
                FolderPermissionLevel.FreeBusyTimeAndSubjectAndLocation, new FolderPermission
                {
                    _canCreateItems = false,
                    _canCreateSubFolders = false,
                    _deleteItems = PermissionScope.None,
                    _editItems = PermissionScope.None,
                    _isFolderContact = false,
                    _isFolderOwner = false,
                    _isFolderVisible = false,
                    _readItems = FolderPermissionReadAccess.TimeAndSubjectAndLocation,
                }
            },
        }
    );

    #endregion


    /// <summary>
    ///     Variants of pre-defined permission levels that Outlook also displays with the same levels.
    /// </summary>
    private static readonly LazyMember<List<FolderPermission>> LevelVariants = new(
        () =>
        {
            var results = new List<FolderPermission>();

            var permissionNone = DefaultPermissions.Member[FolderPermissionLevel.None];
            var permissionOwner = DefaultPermissions.Member[FolderPermissionLevel.Owner];

            // PermissionLevelNoneOption1
            var permission = permissionNone.Clone();
            permission._isFolderVisible = true;
            results.Add(permission);

            // PermissionLevelNoneOption2
            permission = permissionNone.Clone();
            permission._isFolderContact = true;
            results.Add(permission);

            // PermissionLevelNoneOption3
            permission = permissionNone.Clone();
            permission._isFolderContact = true;
            permission._isFolderVisible = true;
            results.Add(permission);

            // PermissionLevelOwnerOption1
            permission = permissionOwner.Clone();
            permission._isFolderContact = false;
            results.Add(permission);

            return results;
        }
    );

    private UserId? _userId;
    private bool _canCreateItems;
    private bool _canCreateSubFolders;
    private bool _isFolderOwner;
    private bool _isFolderVisible;
    private bool _isFolderContact;
    private PermissionScope _editItems;
    private PermissionScope _deleteItems;
    private FolderPermissionReadAccess _readItems;
    private FolderPermissionLevel _permissionLevel;

    /// <summary>
    ///     Determines whether the specified folder permission is the same as this one. The comparison
    ///     does not take UserId and PermissionLevel into consideration.
    /// </summary>
    /// <param name="permission">The folder permission to compare with this folder permission.</param>
    /// <returns>
    ///     True is the specified folder permission is equal to this one, false otherwise.
    /// </returns>
    private bool IsEqualTo(FolderPermission permission)
    {
        return CanCreateItems == permission.CanCreateItems &&
               CanCreateSubFolders == permission.CanCreateSubFolders &&
               IsFolderContact == permission.IsFolderContact &&
               IsFolderVisible == permission.IsFolderVisible &&
               IsFolderOwner == permission.IsFolderOwner &&
               EditItems == permission.EditItems &&
               DeleteItems == permission.DeleteItems &&
               ReadItems == permission.ReadItems;
    }

    /// <summary>
    ///     Create a copy of this FolderPermission instance.
    /// </summary>
    /// <returns>
    ///     Clone of this instance.
    /// </returns>
    private FolderPermission Clone()
    {
        return (FolderPermission)MemberwiseClone();
    }

    /// <summary>
    ///     Determines the permission level of this folder permission based on its individual settings,
    ///     and sets the PermissionLevel property accordingly.
    /// </summary>
    private void AdjustPermissionLevel()
    {
        foreach (var keyValuePair in DefaultPermissions.Member)
        {
            if (IsEqualTo(keyValuePair.Value))
            {
                _permissionLevel = keyValuePair.Key;
                return;
            }
        }

        _permissionLevel = FolderPermissionLevel.Custom;
    }

    /// <summary>
    ///     Copies the values of the individual permissions of the specified folder permission
    ///     to this folder permissions.
    /// </summary>
    /// <param name="permission">The folder permission to copy the values from.</param>
    private void AssignIndividualPermissions(FolderPermission permission)
    {
        _canCreateItems = permission.CanCreateItems;
        _canCreateSubFolders = permission.CanCreateSubFolders;
        _isFolderContact = permission.IsFolderContact;
        _isFolderOwner = permission.IsFolderOwner;
        _isFolderVisible = permission.IsFolderVisible;
        _editItems = permission.EditItems;
        _deleteItems = permission.DeleteItems;
        _readItems = permission.ReadItems;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="FolderPermission" /> class.
    /// </summary>
    public FolderPermission()
    {
        UserId = new UserId();
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="FolderPermission" /> class.
    /// </summary>
    /// <param name="userId">The Id of the user  the permission applies to.</param>
    /// <param name="permissionLevel">The level of the permission.</param>
    public FolderPermission(UserId userId, FolderPermissionLevel permissionLevel)
    {
        EwsUtilities.ValidateParam(userId);

        _userId = userId;
        PermissionLevel = permissionLevel;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="FolderPermission" /> class.
    /// </summary>
    /// <param name="primarySmtpAddress">The primary SMTP address of the user the permission applies to.</param>
    /// <param name="permissionLevel">The level of the permission.</param>
    public FolderPermission(string primarySmtpAddress, FolderPermissionLevel permissionLevel)
    {
        _userId = new UserId(primarySmtpAddress);
        PermissionLevel = permissionLevel;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="FolderPermission" /> class.
    /// </summary>
    /// <param name="standardUser">The standard user the permission applies to.</param>
    /// <param name="permissionLevel">The level of the permission.</param>
    public FolderPermission(StandardUser standardUser, FolderPermissionLevel permissionLevel)
    {
        _userId = new UserId(standardUser);
        PermissionLevel = permissionLevel;
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    /// <param name="isCalendarFolder">if set to <c>true</c> calendar permissions are allowed.</param>
    /// <param name="permissionIndex">Index of the permission.</param>
    internal void Validate(bool isCalendarFolder, int permissionIndex)
    {
        // Check UserId
        if (!UserId.IsValid())
        {
            throw new ServiceValidationException(
                string.Format(Strings.FolderPermissionHasInvalidUserId, permissionIndex)
            );
        }

        // If this permission is to be used for a non-calendar folder make sure that read access and permission level aren't set to Calendar-only values
        if (!isCalendarFolder)
        {
            if (_readItems == FolderPermissionReadAccess.TimeAndSubjectAndLocation ||
                _readItems == FolderPermissionReadAccess.TimeOnly)
            {
                throw new ServiceLocalException(
                    string.Format(Strings.ReadAccessInvalidForNonCalendarFolder, _readItems)
                );
            }

            if (_permissionLevel == FolderPermissionLevel.FreeBusyTimeAndSubjectAndLocation ||
                _permissionLevel == FolderPermissionLevel.FreeBusyTimeOnly)
            {
                throw new ServiceLocalException(
                    string.Format(Strings.PermissionLevelInvalidForNonCalendarFolder, _permissionLevel)
                );
            }
        }
    }

    /// <summary>
    ///     Gets the Id of the user the permission applies to.
    /// </summary>
    public UserId? UserId
    {
        get => _userId;

        set
        {
            if (_userId != null)
            {
                _userId.OnChange -= PropertyChanged;
            }

            SetFieldValue(ref _userId, value);

            if (_userId != null)
            {
                _userId.OnChange += PropertyChanged;
            }
        }
    }

    /// <summary>
    ///     Gets or sets a value indicating whether the user can create new items.
    /// </summary>
    public bool CanCreateItems
    {
        get => _canCreateItems;
        set
        {
            SetFieldValue(ref _canCreateItems, value);
            AdjustPermissionLevel();
        }
    }

    /// <summary>
    ///     Gets or sets a value indicating whether the user can create sub-folders.
    /// </summary>
    public bool CanCreateSubFolders
    {
        get => _canCreateSubFolders;
        set
        {
            SetFieldValue(ref _canCreateSubFolders, value);
            AdjustPermissionLevel();
        }
    }

    /// <summary>
    ///     Gets or sets a value indicating whether the user owns the folder.
    /// </summary>
    public bool IsFolderOwner
    {
        get => _isFolderOwner;
        set
        {
            SetFieldValue(ref _isFolderOwner, value);
            AdjustPermissionLevel();
        }
    }

    /// <summary>
    ///     Gets or sets a value indicating whether the folder is visible to the user.
    /// </summary>
    public bool IsFolderVisible
    {
        get => _isFolderVisible;
        set
        {
            SetFieldValue(ref _isFolderVisible, value);
            AdjustPermissionLevel();
        }
    }

    /// <summary>
    ///     Gets or sets a value indicating whether the user is a contact for the folder.
    /// </summary>
    public bool IsFolderContact
    {
        get => _isFolderContact;
        set
        {
            SetFieldValue(ref _isFolderContact, value);
            AdjustPermissionLevel();
        }
    }

    /// <summary>
    ///     Gets or sets a value indicating if/how the user can edit existing items.
    /// </summary>
    public PermissionScope EditItems
    {
        get => _editItems;
        set
        {
            SetFieldValue(ref _editItems, value);
            AdjustPermissionLevel();
        }
    }

    /// <summary>
    ///     Gets or sets a value indicating if/how the user can delete existing items.
    /// </summary>
    public PermissionScope DeleteItems
    {
        get => _deleteItems;
        set
        {
            SetFieldValue(ref _deleteItems, value);
            AdjustPermissionLevel();
        }
    }

    /// <summary>
    ///     Gets or sets the read items access permission.
    /// </summary>
    public FolderPermissionReadAccess ReadItems
    {
        get => _readItems;
        set
        {
            SetFieldValue(ref _readItems, value);
            AdjustPermissionLevel();
        }
    }

    /// <summary>
    ///     Gets or sets the permission level.
    /// </summary>
    public FolderPermissionLevel PermissionLevel
    {
        get => _permissionLevel;
        set
        {
            if (_permissionLevel != value)
            {
                if (value == FolderPermissionLevel.Custom)
                {
                    throw new ServiceLocalException(Strings.CannotSetPermissionLevelToCustom);
                }

                AssignIndividualPermissions(DefaultPermissions.Member[value]);
                SetFieldValue(ref _permissionLevel, value);
            }
        }
    }

    /// <summary>
    ///     Gets the permission level that Outlook would display for this folder permission.
    /// </summary>
    public FolderPermissionLevel DisplayPermissionLevel
    {
        get
        {
            // If permission level is set to Custom, see if there's a variant
            // that Outlook would map to the same permission level.
            if (_permissionLevel == FolderPermissionLevel.Custom)
            {
                foreach (var variant in LevelVariants.Member)
                {
                    if (IsEqualTo(variant))
                    {
                        return variant.PermissionLevel;
                    }
                }
            }

            return _permissionLevel;
        }
    }

    /// <summary>
    ///     Property was changed.
    /// </summary>
    /// <param name="complexProperty">The complex property.</param>
    private void PropertyChanged(ComplexProperty complexProperty)
    {
        Changed();
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.UserId:
            {
                UserId = new UserId();
                UserId.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.CanCreateItems:
            {
                _canCreateItems = reader.ReadValue<bool>();
                return true;
            }
            case XmlElementNames.CanCreateSubFolders:
            {
                _canCreateSubFolders = reader.ReadValue<bool>();
                return true;
            }
            case XmlElementNames.IsFolderOwner:
            {
                _isFolderOwner = reader.ReadValue<bool>();
                return true;
            }
            case XmlElementNames.IsFolderVisible:
            {
                _isFolderVisible = reader.ReadValue<bool>();
                return true;
            }
            case XmlElementNames.IsFolderContact:
            {
                _isFolderContact = reader.ReadValue<bool>();
                return true;
            }
            case XmlElementNames.EditItems:
            {
                _editItems = reader.ReadValue<PermissionScope>();
                return true;
            }
            case XmlElementNames.DeleteItems:
            {
                _deleteItems = reader.ReadValue<PermissionScope>();
                return true;
            }
            case XmlElementNames.ReadItems:
            {
                _readItems = reader.ReadValue<FolderPermissionReadAccess>();
                return true;
            }
            case XmlElementNames.PermissionLevel:
            case XmlElementNames.CalendarPermissionLevel:
            {
                _permissionLevel = reader.ReadValue<FolderPermissionLevel>();
                return true;
            }
            default:
            {
                return false;
            }
        }
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal override void LoadFromXml(EwsServiceXmlReader reader, XmlNamespace xmlNamespace, string xmlElementName)
    {
        base.LoadFromXml(reader, xmlNamespace, xmlElementName);

        AdjustPermissionLevel();
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="isCalendarFolder">If true, this permission is for a calendar folder.</param>
    internal void WriteElementsToXml(EwsServiceXmlWriter writer, bool isCalendarFolder)
    {
        if (UserId != null)
        {
            UserId.WriteToXml(writer, XmlElementNames.UserId);
        }

        if (PermissionLevel == FolderPermissionLevel.Custom)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.CanCreateItems, CanCreateItems);

            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.CanCreateSubFolders, CanCreateSubFolders);

            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsFolderOwner, IsFolderOwner);

            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsFolderVisible, IsFolderVisible);

            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsFolderContact, IsFolderContact);

            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.EditItems, EditItems);

            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.DeleteItems, DeleteItems);

            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ReadItems, ReadItems);
        }

        writer.WriteElementValue(
            XmlNamespace.Types,
            isCalendarFolder ? XmlElementNames.CalendarPermissionLevel : XmlElementNames.PermissionLevel,
            PermissionLevel
        );
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <param name="isCalendarFolder">If true, this permission is for a calendar folder.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName, bool isCalendarFolder)
    {
        writer.WriteStartElement(Namespace, xmlElementName);
        WriteAttributesToXml(writer);
        WriteElementsToXml(writer, isCalendarFolder);
        writer.WriteEndElement();
    }
}
