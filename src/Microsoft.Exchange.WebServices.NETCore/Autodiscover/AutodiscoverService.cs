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
using System.Net;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;

using JetBrains.Annotations;

using Microsoft.Exchange.WebServices.Data;

namespace Microsoft.Exchange.WebServices.Autodiscover;

/// <summary>
///     Defines a delegate that is used by the AutodiscoverService to ask whether a redirectionUrl can be used.
/// </summary>
/// <param name="redirectionUrl">Redirection URL that Autodiscover wants to use.</param>
/// <returns>Delegate returns true if Autodiscover is allowed to use this URL.</returns>
public delegate bool AutodiscoverRedirectionUrlValidationCallback(string redirectionUrl);

/// <summary>
///     Represents a binding to the Exchange Autodiscover Service.
/// </summary>
[PublicAPI]
public sealed class AutodiscoverService : ExchangeServiceBase
{
    #region Static members

    private static readonly TimeSpan AutodiscoverTimeout = TimeSpan.FromSeconds(10);

    /// <summary>
    ///     Autodiscover legacy path
    /// </summary>
    private const string AutodiscoverLegacyPath = "/autodiscover/autodiscover.xml";

    /// <summary>
    ///     Autodiscover legacy Url with protocol fill-in
    /// </summary>
    private const string AutodiscoverLegacyUrl = "{0}://{1}" + AutodiscoverLegacyPath;

    /// <summary>
    ///     Autodiscover legacy HTTPS Url
    /// </summary>
    private const string AutodiscoverLegacyHttpsUrl = "https://{0}" + AutodiscoverLegacyPath;

    /// <summary>
    ///     Autodiscover legacy HTTP Url
    /// </summary>
    private const string AutodiscoverLegacyHttpUrl = "http://{0}" + AutodiscoverLegacyPath;

    /// <summary>
    ///     Autodiscover SOAP HTTPS Url
    /// </summary>
    private const string AutodiscoverSoapHttpsUrl = "https://{0}/autodiscover/autodiscover.svc";

    /// <summary>
    ///     Autodiscover SOAP WS-Security HTTPS Url
    /// </summary>
    private const string AutodiscoverSoapWsSecurityHttpsUrl = AutodiscoverSoapHttpsUrl + "/wssecurity";

    /// <summary>
    ///     Autodiscover SOAP WS-Security symmetrickey HTTPS Url
    /// </summary>
    private const string AutodiscoverSoapWsSecuritySymmetricKeyHttpsUrl =
        AutodiscoverSoapHttpsUrl + "/wssecurity/symmetrickey";

    /// <summary>
    ///     Autodiscover SOAP WS-Security x509cert HTTPS Url
    /// </summary>
    private const string AutodiscoverSoapWsSecurityX509CertHttpsUrl = AutodiscoverSoapHttpsUrl + "/wssecurity/x509cert";

    /// <summary>
    ///     Autodiscover request namespace
    /// </summary>
    private const string AutodiscoverRequestNamespace =
        "http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006";

    /// <summary>
    ///     Legacy path regular expression.
    /// </summary>
    private static readonly Regex LegacyPathRegex = new Regex(
        @"/autodiscover/([^/]+/)*autodiscover.xml",
        RegexOptions.Compiled | RegexOptions.IgnoreCase
    );

    /// <summary>
    ///     Maximum number of Url (or address) redirections that will be followed by an Autodiscover call
    /// </summary>
    internal const int AutodiscoverMaxRedirections = 10;

    /// <summary>
    ///     HTTP header indicating that SOAP Autodiscover service is enabled.
    /// </summary>
    private const string AutodiscoverSoapEnabledHeaderName = "X-SOAP-Enabled";

    /// <summary>
    ///     HTTP header indicating that WS-Security Autodiscover service is enabled.
    /// </summary>
    private const string AutodiscoverWsSecurityEnabledHeaderName = "X-WSSecurity-Enabled";

    /// <summary>
    ///     HTTP header indicating that WS-Security/SymmetricKey Autodiscover service is enabled.
    /// </summary>
    private const string AutodiscoverWsSecuritySymmetricKeyEnabledHeaderName = "X-WSSecurity-SymmetricKey-Enabled";

    /// <summary>
    ///     HTTP header indicating that WS-Security/X509Cert Autodiscover service is enabled.
    /// </summary>
    private const string AutodiscoverWsSecurityX509CertEnabledHeaderName = "X-WSSecurity-X509Cert-Enabled";

    /// <summary>
    ///     HTTP header indicating that OAuth Autodiscover service is enabled.
    /// </summary>
    private const string AutodiscoverOAuthEnabledHeaderName = "X-OAuth-Enabled";

    /// <summary>
    ///     Minimum request version for Autodiscover SOAP service.
    /// </summary>
    private const ExchangeVersion MinimumRequestVersionForAutoDiscoverSoapService = ExchangeVersion.Exchange2010;

    #endregion


    #region Private members

    private string? _domain;
    private Uri? _url;
    private readonly AutodiscoverDnsClient _dnsClient;

    private delegate Task<Tuple<TGetSettingsResponseCollection, Uri>>
        GetSettingsMethod<TGetSettingsResponseCollection, TSettingName>(
            List<string> smtpAddresses,
            List<TSettingName> settings,
            ExchangeVersion? requestedVersion,
            Uri autodiscoverUrl
        );

    #endregion


    /// <summary>
    ///     Default implementation of AutodiscoverRedirectionUrlValidationCallback.
    ///     Always returns true indicating that the URL can be used.
    /// </summary>
    /// <param name="redirectionUrl">The redirection URL.</param>
    /// <returns>Returns true.</returns>
    private bool DefaultAutodiscoverRedirectionUrlValidationCallback(string redirectionUrl)
    {
        throw new AutodiscoverLocalException(string.Format(Strings.AutodiscoverRedirectBlocked, redirectionUrl));
    }


    #region Legacy Autodiscover

    /// <summary>
    ///     Calls the Autodiscover service to get configuration settings at the specified URL.
    /// </summary>
    /// <typeparam name="TSettings">The type of the settings to retrieve.</typeparam>
    /// <param name="emailAddress">The email address to retrieve configuration settings for.</param>
    /// <param name="url">The URL of the Autodiscover service.</param>
    /// <returns>The requested configuration settings.</returns>
    private async Task<TSettings> GetLegacyUserSettingsAtUrl<TSettings>(string emailAddress, Uri url)
        where TSettings : ConfigurationSettingsBase, new()
    {
        TraceMessage(TraceFlags.AutodiscoverConfiguration, $"Trying to call Autodiscover for {emailAddress} on {url}.");

        var settings = new TSettings();

        var request = PrepareHttpRequestMessageForUrl(url);

        using (var requestStream = new MemoryStream())
        {
            // If tracing is enabled, we generate the request in-memory so that we
            // can pass it along to the ITraceListener. Then we copy the stream to
            // the request stream.
            if (IsTraceEnabledFor(TraceFlags.AutodiscoverRequest))
            {
                using var memoryStream = new MemoryStream();
                await using var writer = new StreamWriter(memoryStream);
                WriteLegacyAutodiscoverRequest(emailAddress, settings, writer);
                await writer.FlushAsync();

                TraceXml(TraceFlags.AutodiscoverRequest, memoryStream);

                EwsUtilities.CopyStream(memoryStream, requestStream);
            }
            else
            {
                await using var writer = new StreamWriter(requestStream);
                WriteLegacyAutodiscoverRequest(emailAddress, settings, writer);
            }

            request.Content = new ByteArrayContent(requestStream.ToArray());
            request.Content.Headers.ContentType = new MediaTypeHeaderValue("text/xml")
            {
                CharSet = "utf-8",
            };
        }

        using var client = PrepareHttpClient();
        using IEwsHttpWebResponse webResponse = new EwsHttpWebResponse(client.SendAsync(request).Result);
        if (TryGetRedirectionResponse(webResponse, out var redirectUrl))
        {
            settings.MakeRedirectionResponse(redirectUrl);
            return settings;
        }

        await using (var responseStream = await webResponse.GetResponseStream())
        {
            // If tracing is enabled, we read the entire response into a MemoryStream so that we
            // can pass it along to the ITraceListener. Then we parse the response from the 
            // MemoryStream.
            if (IsTraceEnabledFor(TraceFlags.AutodiscoverResponse))
            {
                using var memoryStream = new MemoryStream();
                // Copy response stream to in-memory stream and reset to start
                EwsUtilities.CopyStream(responseStream, memoryStream);
                memoryStream.Position = 0;

                TraceResponse(webResponse, memoryStream);

                var reader = new EwsXmlReader(memoryStream);
                reader.Read(XmlNodeType.XmlDeclaration);
                settings.LoadFromXml(reader);
            }
            else
            {
                var reader = new EwsXmlReader(responseStream);
                reader.Read(XmlNodeType.XmlDeclaration);
                settings.LoadFromXml(reader);
            }
        }

        return settings;
    }

    /// <summary>
    ///     Writes the autodiscover request.
    /// </summary>
    /// <param name="emailAddress">The email address.</param>
    /// <param name="settings">The settings.</param>
    /// <param name="writer">The writer.</param>
    private static void WriteLegacyAutodiscoverRequest(
        string emailAddress,
        ConfigurationSettingsBase settings,
        StreamWriter writer
    )
    {
        writer.Write("<Autodiscover xmlns=\"{0}\">", AutodiscoverRequestNamespace);
        writer.Write("<Request>");
        writer.Write("<EMailAddress>{0}</EMailAddress>", emailAddress);
        writer.Write("<AcceptableResponseSchema>{0}</AcceptableResponseSchema>", settings.GetNamespace());
        writer.Write("</Request>");
        writer.Write("</Autodiscover>");
    }

    /// <summary>
    ///     Gets a redirection URL to an SSL-enabled Autodiscover service from the standard non-SSL Autodiscover URL.
    /// </summary>
    /// <param name="domainName">The name of the domain to call Autodiscover on.</param>
    /// <returns>A valid SSL-enabled redirection URL. (May be null).</returns>
    private Uri? GetRedirectUrl(string domainName)
    {
        var url = string.Format(AutodiscoverLegacyHttpUrl, "autodiscover." + domainName);

        TraceMessage(TraceFlags.AutodiscoverConfiguration, $"Trying to get Autodiscover redirection URL from {url}.");

        IEwsHttpWebResponse? response = null;

        try
        {
            using var client = new HttpClient(
                new HttpClientHandler
                {
                    AllowAutoRedirect = false,
                }
            );
            client.Timeout = AutodiscoverTimeout;
            var httpResponse = client.GetAsync(url).Result;
            response = new EwsHttpWebResponse(httpResponse);
        }
        catch (Exception ex)
        {
            TraceMessage(TraceFlags.AutodiscoverConfiguration, $"I/O error: {ex.Message}");
        }

        if (response != null)
        {
            using (response)
            {
                if (TryGetRedirectionResponse(response, out var redirectUrl))
                {
                    return redirectUrl;
                }
            }
        }

        TraceMessage(TraceFlags.AutodiscoverConfiguration, "No Autodiscover redirection URL was returned.");

        return null;
    }

    /// <summary>
    ///     Tries the get redirection response.
    /// </summary>
    /// <param name="response">The response.</param>
    /// <param name="redirectUrl">The redirect URL.</param>
    /// <returns>True if a valid redirection URL was found.</returns>
    private bool TryGetRedirectionResponse(IEwsHttpWebResponse response, [MaybeNullWhen(false)] out Uri redirectUrl)
    {
        redirectUrl = null;
        if (AutodiscoverRequest.IsRedirectionResponse(response))
        {
            // Get the redirect location and verify that it's valid.
            var location = response.Headers.Location;

            if (location != null)
            {
                try
                {
                    redirectUrl = new Uri(response.ResponseUri, location);

                    // Check if URL is SSL and that the path matches.
                    var match = LegacyPathRegex.Match(redirectUrl.AbsolutePath);
                    if (redirectUrl.Scheme == "https" && match.Success)
                    {
                        TraceMessage(TraceFlags.AutodiscoverConfiguration, $"Redirection URL found: '{redirectUrl}'");

                        return true;
                    }
                }
                catch (UriFormatException)
                {
                    TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        $"Invalid redirection URL was returned: '{location}'"
                    );
                    return false;
                }
            }
        }

        return false;
    }

    /// <summary>
    ///     Calls the legacy Autodiscover service to retrieve configuration settings.
    /// </summary>
    /// <typeparam name="TSettings">The type of the settings to retrieve.</typeparam>
    /// <param name="emailAddress">The email address to retrieve configuration settings for.</param>
    /// <returns>The requested configuration settings.</returns>
    internal async Task<TSettings> GetLegacyUserSettings<TSettings>(string emailAddress)
        where TSettings : ConfigurationSettingsBase, new()
    {
        // If Url is specified, call service directly.
        if (Url != null)
        {
            var match = LegacyPathRegex.Match(Url.AbsolutePath);
            if (match.Success)
            {
                return await GetLegacyUserSettingsAtUrl<TSettings>(emailAddress, Url);
            }

            // this.Uri is intended for Autodiscover SOAP service, convert to Legacy endpoint URL.
            var autodiscoverUrl = new Uri(Url, AutodiscoverLegacyPath);
            return await GetLegacyUserSettingsAtUrl<TSettings>(emailAddress, autodiscoverUrl);
        }

        // If Domain is specified, figure out the endpoint Url and call service.

        if (!string.IsNullOrEmpty(Domain))
        {
            var autodiscoverUrl = new Uri(string.Format(AutodiscoverLegacyHttpsUrl, Domain));
            return await GetLegacyUserSettingsAtUrl<TSettings>(emailAddress, autodiscoverUrl);
        }

        // No Url or Domain specified, need to figure out which endpoint to use.
        const int currentHop = 1;
        var redirectionEmailAddresses = new List<string>();
        return (await InternalGetLegacyUserSettings<TSettings>(emailAddress, redirectionEmailAddresses, currentHop))
            .Item1;
    }

    /// <summary>
    ///     Calls the legacy Autodiscover service to retrieve configuration settings.
    /// </summary>
    /// <typeparam name="TSettings">The type of the settings to retrieve.</typeparam>
    /// <param name="emailAddress">The email address to retrieve configuration settings for.</param>
    /// <param name="redirectionEmailAddresses">List of previous email addresses.</param>
    /// <param name="currentHop">Current number of redirection urls/addresses attempted so far.</param>
    /// <returns>The requested configuration settings.</returns>
    private async Task<Tuple<TSettings, int>> InternalGetLegacyUserSettings<TSettings>(
        string emailAddress,
        List<string> redirectionEmailAddresses,
        int currentHop
    )
        where TSettings : ConfigurationSettingsBase, new()
    {
        var domainName = EwsUtilities.DomainFromEmailAddress(emailAddress);

        var urls = GetAutodiscoverServiceUrls(domainName, out var scpUrlCount);

        if (urls.Count == 0)
        {
            throw new ServiceValidationException(Strings.AutodiscoverServiceRequestRequiresDomainOrUrl);
        }

        // Assume caller is not inside the Intranet, regardless of whether SCP Urls 
        // were returned or not. SCP Urls are only relevant if one of them returns
        // valid Autodiscover settings.
        IsExternal = true;

        var currentUrlIndex = 0;

        // Used to save exception for later reporting.
        Exception? delayedException = null;
        TSettings? settings = null;

        do
        {
            var autodiscoverUrl = urls[currentUrlIndex];
            var isScpUrl = currentUrlIndex < scpUrlCount;

            try
            {
                settings = await GetLegacyUserSettingsAtUrl<TSettings>(emailAddress, autodiscoverUrl);

                switch (settings.ResponseType)
                {
                    case AutodiscoverResponseType.Success:
                    {
                        // Not external if Autodiscover endpoint found via SCP returned the settings.
                        if (isScpUrl)
                        {
                            IsExternal = false;
                        }

                        Url = autodiscoverUrl;
                        return Tuple.Create(settings, currentHop);
                    }
                    case AutodiscoverResponseType.RedirectUrl:
                    {
                        if (currentHop < AutodiscoverMaxRedirections)
                        {
                            currentHop++;
                            TraceMessage(
                                TraceFlags.AutodiscoverResponse,
                                $"Autodiscover service returned redirection URL '{settings.RedirectTarget}'."
                            );

                            urls[currentUrlIndex] = new Uri(settings.RedirectTarget);
                            break;
                        }

                        throw new AutodiscoverLocalException(Strings.MaximumRedirectionHopsExceeded);
                    }
                    case AutodiscoverResponseType.RedirectAddress:
                    {
                        if (currentHop < AutodiscoverMaxRedirections)
                        {
                            currentHop++;
                            TraceMessage(
                                TraceFlags.AutodiscoverResponse,
                                $"Autodiscover service returned redirection email address '{settings.RedirectTarget}'."
                            );

                            // If this email address was already tried, we may have a loop
                            // in SCP lookups. Disable consideration of SCP records.
                            DisableScpLookupIfDuplicateRedirection(settings.RedirectTarget, redirectionEmailAddresses);

                            return await InternalGetLegacyUserSettings<TSettings>(
                                settings.RedirectTarget,
                                redirectionEmailAddresses,
                                currentHop
                            );
                        }

                        throw new AutodiscoverLocalException(Strings.MaximumRedirectionHopsExceeded);
                    }
                    case AutodiscoverResponseType.Error:
                    {
                        // Don't treat errors from an SCP-based Autodiscover service to be conclusive.
                        // We'll try the next one and record the error for later.
                        if (isScpUrl)
                        {
                            TraceMessage(
                                TraceFlags.AutodiscoverConfiguration,
                                "Error returned by Autodiscover service found via SCP, treating as inconclusive."
                            );

                            delayedException = new AutodiscoverRemoteException(
                                Strings.AutodiscoverError,
                                settings.Error
                            );
                            currentUrlIndex++;
                        }
                        else
                        {
                            throw new AutodiscoverRemoteException(Strings.AutodiscoverError, settings.Error);
                        }

                        break;
                    }
                    default:
                    {
                        EwsUtilities.Assert(
                            false,
                            "Autodiscover.GetConfigurationSettings",
                            "An unexpected error has occurred. This code path should never be reached."
                        );
                        break;
                    }
                }
            }
            catch (EwsHttpClientException ex)
            {
                if (ex.Response != null)
                {
                    var response = HttpWebRequestFactory.CreateExceptionResponse(ex);
                    if (TryGetRedirectionResponse(response, out var redirectUrl))
                    {
                        TraceMessage(
                            TraceFlags.AutodiscoverConfiguration,
                            $"Host returned a redirection to url {redirectUrl}"
                        );

                        currentHop++;
                        urls[currentUrlIndex] = redirectUrl;
                    }
                    else
                    {
                        ProcessHttpErrorResponse(response, ex);

                        TraceMessage(
                            TraceFlags.AutodiscoverConfiguration,
                            $"{_url} failed: {ex.GetType().Name} ({ex.Message})"
                        );

                        // The url did not work, let's try the next.
                        currentUrlIndex++;
                    }
                }
                else
                {
                    TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        $"{_url} failed: {ex.GetType().Name} ({ex.Message})"
                    );

                    // The url did not work, let's try the next.
                    currentUrlIndex++;
                }
            }
            catch (XmlException ex)
            {
                TraceMessage(TraceFlags.AutodiscoverConfiguration, $"{_url} failed: XML parsing error: {ex.Message}");

                // The content at the URL wasn't a valid response, let's try the next.
                currentUrlIndex++;
            }
            catch (Exception ex)
            {
                TraceMessage(TraceFlags.AutodiscoverConfiguration, $"{_url} failed: I/O error: {ex.Message}");

                // The content at the URL wasn't a valid response, let's try the next.
                currentUrlIndex++;
            }
        } while (currentUrlIndex < urls.Count);

        // If we got this far it's because none of the URLs we tried have worked. As a next-to-last chance, use GetRedirectUrl to 
        // try to get a redirection URL using an HTTP GET on a non-SSL Autodiscover endpoint. If successful, use this 
        // redirection URL to get the configuration settings for this email address. (This will be a common scenario for 
        // DataCenter deployments).
        var redirectionUrl = GetRedirectUrl(domainName);
        if (redirectionUrl != null)
        {
            var result = await TryLastChanceHostRedirection<TSettings>(emailAddress, redirectionUrl);
            if (result.Item1)
            {
                return Tuple.Create(settings, currentHop);
            }
        }

        {
            // Getting a redirection URL from an HTTP GET failed too. As a last chance, try to get an appropriate SRV Record
            // using DnsQuery. If successful, use this redirection URL to get the configuration settings for this email address.
            redirectionUrl = GetRedirectionUrlFromDnsSrvRecord(domainName);
            if (redirectionUrl != null)
            {
                var result = await TryLastChanceHostRedirection<TSettings>(emailAddress, redirectionUrl);
                if (result.Item1)
                {
                    return Tuple.Create(settings, currentHop);
                }
            }

            // If there was an earlier exception, throw it.
            if (delayedException != null)
            {
                throw delayedException;
            }

            throw new AutodiscoverLocalException(Strings.AutodiscoverCouldNotBeLocated);
        }
    }

    /// <summary>
    ///     Get an autodiscover SRV record in DNS and construct autodiscover URL.
    /// </summary>
    /// <param name="domainName">Name of the domain.</param>
    /// <returns>Autodiscover URL (may be null if lookup failed)</returns>
    internal Uri? GetRedirectionUrlFromDnsSrvRecord(string domainName)
    {
        TraceMessage(
            TraceFlags.AutodiscoverConfiguration,
            $"Trying to get Autodiscover host from DNS SRV record for {domainName}."
        );

        var hostname = _dnsClient.FindAutodiscoverHostFromSrv(domainName);
        if (!string.IsNullOrEmpty(hostname))
        {
            TraceMessage(TraceFlags.AutodiscoverConfiguration, $"Autodiscover host {hostname} was returned.");

            return new Uri(string.Format(AutodiscoverLegacyHttpsUrl, hostname));
        }

        TraceMessage(TraceFlags.AutodiscoverConfiguration, "No matching Autodiscover DNS SRV records were found.");

        return null;
    }

    /// <summary>
    ///     Tries to get Autodiscover settings using redirection Url.
    /// </summary>
    /// <typeparam name="TSettings">The type of the settings.</typeparam>
    /// <param name="emailAddress">The email address.</param>
    /// <param name="redirectionUrl">Redirection Url.</param>
    private async Task<Tuple<bool, TSettings>> TryLastChanceHostRedirection<TSettings>(
        string emailAddress,
        Uri redirectionUrl
    )
        where TSettings : ConfigurationSettingsBase, new()
    {
        TSettings? settings = null;

        var redirectionEmailAddresses = new List<string>();

        // Bug 60274: Performing a non-SSL HTTP GET to retrieve a redirection URL is potentially unsafe. We allow the caller 
        // to specify delegate to be called to determine whether we are allowed to use the redirection URL. 
        if (CallRedirectionUrlValidationCallback(redirectionUrl.ToString()))
        {
            for (var currentHop = 0; currentHop < AutodiscoverMaxRedirections; currentHop++)
            {
                try
                {
                    settings = await GetLegacyUserSettingsAtUrl<TSettings>(emailAddress, redirectionUrl);

                    switch (settings.ResponseType)
                    {
                        case AutodiscoverResponseType.Success:
                        {
                            return Tuple.Create(true, settings);
                        }
                        case AutodiscoverResponseType.Error:
                        {
                            throw new AutodiscoverRemoteException(Strings.AutodiscoverError, settings.Error);
                        }
                        case AutodiscoverResponseType.RedirectAddress:
                        {
                            // If this email address was already tried, we may have a loop
                            // in SCP lookups. Disable consideration of SCP records.
                            DisableScpLookupIfDuplicateRedirection(settings.RedirectTarget, redirectionEmailAddresses);

                            var result = await InternalGetLegacyUserSettings<TSettings>(
                                settings.RedirectTarget,
                                redirectionEmailAddresses,
                                currentHop
                            );
                            settings = result.Item1;
                            currentHop = result.Item2;
                            return Tuple.Create(true, settings);
                        }

                        case AutodiscoverResponseType.RedirectUrl:
                        {
                            try
                            {
                                redirectionUrl = new Uri(settings.RedirectTarget);
                            }
                            catch (UriFormatException)
                            {
                                TraceMessage(
                                    TraceFlags.AutodiscoverConfiguration,
                                    $"Service returned invalid redirection URL {settings.RedirectTarget}"
                                );
                                return Tuple.Create(false, settings);
                            }

                            break;
                        }
                        default:
                        {
                            var failureMessage =
                                $"Autodiscover call at {redirectionUrl} failed with error {settings.ResponseType}, target {settings.RedirectTarget}";
                            TraceMessage(TraceFlags.AutodiscoverConfiguration, failureMessage);
                            return Tuple.Create(false, settings);
                        }
                    }
                }
                catch (EwsHttpClientException ex)
                {
                    if (ex.Response != null)
                    {
                        var response = HttpWebRequestFactory.CreateExceptionResponse(ex);
                        if (TryGetRedirectionResponse(response, out redirectionUrl))
                        {
                            TraceMessage(
                                TraceFlags.AutodiscoverConfiguration,
                                $"Host returned a redirection to url {redirectionUrl}"
                            );
                            continue;
                        }

                        ProcessHttpErrorResponse(response, ex);
                    }

                    TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        $"{_url} failed: {ex.GetType().Name} ({ex.Message})"
                    );

                    return Tuple.Create(false, settings);
                }
                catch (XmlException ex)
                {
                    // If the response is malformed, it wasn't a valid Autodiscover endpoint.
                    TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        $"{redirectionUrl} failed: XML parsing error: {ex.Message}"
                    );

                    return Tuple.Create(false, settings);
                }
                catch (Exception ex)
                {
                    TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        $"{redirectionUrl} failed: I/O error: {ex.Message}"
                    );

                    return Tuple.Create(false, settings);
                }
            }
        }

        return Tuple.Create(false, settings);
    }

    /// <summary>
    ///     Disables SCP lookup if duplicate email address redirection.
    /// </summary>
    /// <param name="emailAddress">The email address to use.</param>
    /// <param name="redirectionEmailAddresses">The list of prior redirection email addresses.</param>
    private void DisableScpLookupIfDuplicateRedirection(string emailAddress, List<string> redirectionEmailAddresses)
    {
        // SMTP addresses are case-insensitive so entries are converted to lower-case.
        emailAddress = emailAddress.ToLowerInvariant();

        if (redirectionEmailAddresses.Contains(emailAddress))
        {
            EnableScpLookup = false;
        }
        else
        {
            redirectionEmailAddresses.Add(emailAddress);
        }
    }

    /// <summary>
    ///     Gets user settings from Autodiscover legacy endpoint.
    /// </summary>
    /// <param name="emailAddress">The email address.</param>
    /// <param name="requestedSettings">The requested settings.</param>
    /// <returns>GetUserSettingsResponse</returns>
    internal async Task<GetUserSettingsResponse> InternalGetLegacyUserSettings(
        string emailAddress,
        List<UserSettingName> requestedSettings
    )
    {
        // Cannot call legacy Autodiscover service with WindowsLive and other WSSecurity-based credentials
        if (Credentials != null && Credentials is WSSecurityBasedCredentials)
        {
            throw new AutodiscoverLocalException(Strings.WLIDCredentialsCannotBeUsedWithLegacyAutodiscover);
        }

        var settings = await GetLegacyUserSettings<OutlookConfigurationSettings>(emailAddress);

        return settings.ConvertSettings(emailAddress, requestedSettings);
    }

    #endregion


    #region SOAP-based Autodiscover

    /// <summary>
    ///     Calls the SOAP Autodiscover service for user settings for a single SMTP address.
    /// </summary>
    /// <param name="smtpAddress">SMTP address.</param>
    /// <param name="requestedSettings">The requested settings.</param>
    /// <returns></returns>
    internal async Task<GetUserSettingsResponse> InternalGetSoapUserSettings(
        string smtpAddress,
        List<UserSettingName> requestedSettings
    )
    {
        var smtpAddresses = new List<string>
        {
            smtpAddress,
        };

        var redirectionEmailAddresses = new List<string>
        {
            smtpAddress.ToLowerInvariant(),
        };

        for (var currentHop = 0; currentHop < AutodiscoverMaxRedirections; currentHop++)
        {
            var response = (await GetUserSettings(smtpAddresses, requestedSettings))[0];

            switch (response.ErrorCode)
            {
                case AutodiscoverErrorCode.RedirectAddress:
                {
                    TraceMessage(
                        TraceFlags.AutodiscoverResponse,
                        $"Autodiscover service returned redirection email address '{response.RedirectTarget}'."
                    );

                    smtpAddresses.Clear();
                    smtpAddresses.Add(response.RedirectTarget.ToLowerInvariant());
                    Url = null;
                    Domain = null;

                    // If this email address was already tried, we may have a loop
                    // in SCP lookups. Disable consideration of SCP records.
                    DisableScpLookupIfDuplicateRedirection(response.RedirectTarget, redirectionEmailAddresses);
                    break;
                }

                case AutodiscoverErrorCode.RedirectUrl:
                {
                    TraceMessage(
                        TraceFlags.AutodiscoverResponse,
                        $"Autodiscover service returned redirection URL '{response.RedirectTarget}'."
                    );

                    Url = Credentials.AdjustUrl(new Uri(response.RedirectTarget));
                    break;
                }

                case AutodiscoverErrorCode.NoError:
                default:
                {
                    return response;
                }
            }
        }

        throw new AutodiscoverLocalException(Strings.AutodiscoverCouldNotBeLocated);
    }

    /// <summary>
    ///     Gets the user settings using Autodiscover SOAP service.
    /// </summary>
    /// <param name="smtpAddresses">The SMTP addresses of the users.</param>
    /// <param name="settings">The settings.</param>
    /// <returns></returns>
    internal Task<GetUserSettingsResponseCollection> GetUserSettings(
        List<string> smtpAddresses,
        List<UserSettingName> settings
    )
    {
        EwsUtilities.ValidateParam(smtpAddresses);
        EwsUtilities.ValidateParam(settings);

        return GetSettings(
            smtpAddresses,
            settings,
            null,
            InternalGetUserSettings,
            () => EwsUtilities.DomainFromEmailAddress(smtpAddresses[0])
        );
    }

    /// <summary>
    ///     Gets user or domain settings using Autodiscover SOAP service.
    /// </summary>
    /// <typeparam name="TGetSettingsResponseCollection">Type of response collection to return.</typeparam>
    /// <typeparam name="TSettingName">Type of setting name.</typeparam>
    /// <param name="identities">Either the domains or the SMTP addresses of the users.</param>
    /// <param name="settings">The settings.</param>
    /// <param name="requestedVersion">Requested version of the Exchange service.</param>
    /// <param name="getSettingsMethod">The method to use.</param>
    /// <param name="getDomainMethod">The method to calculate the domain value.</param>
    /// <returns></returns>
    private async Task<TGetSettingsResponseCollection> GetSettings<TGetSettingsResponseCollection, TSettingName>(
        List<string> identities,
        List<TSettingName> settings,
        ExchangeVersion? requestedVersion,
        GetSettingsMethod<TGetSettingsResponseCollection, TSettingName> getSettingsMethod,
        Func<string> getDomainMethod
    )
    {
        TGetSettingsResponseCollection response;

        // Autodiscover service only exists in E14 or later.
        if (RequestedServerVersion < MinimumRequestVersionForAutoDiscoverSoapService)
        {
            throw new ServiceVersionException(
                string.Format(
                    Strings.AutodiscoverServiceIncompatibleWithRequestVersion,
                    MinimumRequestVersionForAutoDiscoverSoapService
                )
            );
        }

        // If Url is specified, call service directly.
        if (Url != null)
        {
            var autodiscoverUrl = Url;

            var result = await getSettingsMethod(identities, settings, requestedVersion, autodiscoverUrl);
            response = result.Item1;
            autodiscoverUrl = result.Item2;

            Url = autodiscoverUrl;
            return response;
        }

        // If Domain is specified, determine endpoint Url and call service.

        if (!string.IsNullOrEmpty(Domain))
        {
            var autodiscoverUrl = GetAutodiscoverEndpointUrl(Domain);
            var result = await getSettingsMethod(identities, settings, requestedVersion, autodiscoverUrl);
            response = result.Item1;
            autodiscoverUrl = result.Item2;

            // If we got this far, response was successful, set Url.
            Url = autodiscoverUrl;
            return response;
        }

        // No Url or Domain specified, need to figure out which endpoint(s) to try.
        else
        {
            // Assume caller is not inside the Intranet, regardless of whether SCP Urls 
            // were returned or not. SCP Urls are only relevent if one of them returns
            // valid Autodiscover settings.
            IsExternal = true;

            Uri autodiscoverUrl;

            var domainName = getDomainMethod();
            var hosts = GetAutodiscoverServiceHosts(domainName, out var scpHostCount);

            if (hosts.Count == 0)
            {
                throw new ServiceValidationException(Strings.AutodiscoverServiceRequestRequiresDomainOrUrl);
            }

            for (var currentHostIndex = 0; currentHostIndex < hosts.Count; currentHostIndex++)
            {
                var host = hosts[currentHostIndex];
                var isScpHost = currentHostIndex < scpHostCount;

                if (TryGetAutodiscoverEndpointUrl(host, out autodiscoverUrl))
                {
                    try
                    {
                        var result = await getSettingsMethod(identities, settings, requestedVersion, autodiscoverUrl);
                        response = result.Item1;
                        autodiscoverUrl = result.Item2;

                        // If we got this far, the response was successful, set Url.
                        Url = autodiscoverUrl;

                        // Not external if Autodiscover endpoint found via SCP returned the settings.
                        if (isScpHost)
                        {
                            IsExternal = false;
                        }

                        return response;
                    }
                    catch (AutodiscoverResponseException)
                    {
                        // skip
                    }
                    catch (ServiceRequestException)
                    {
                        // skip
                    }
                }
            }

            // Next-to-last chance: try unauthenticated GET over HTTP to be redirected to appropriate service endpoint.
            autodiscoverUrl = GetRedirectUrl(domainName);
            if (autodiscoverUrl != null &&
                CallRedirectionUrlValidationCallback(autodiscoverUrl.ToString()) &&
                TryGetAutodiscoverEndpointUrl(autodiscoverUrl.Host, out autodiscoverUrl))
            {
                var result = await getSettingsMethod(identities, settings, requestedVersion, autodiscoverUrl);
                response = result.Item1;
                autodiscoverUrl = result.Item2;

                // If we got this far, the response was successful, set Url.
                Url = autodiscoverUrl;

                return response;
            }

            // Last Chance: try to read autodiscover SRV Record from DNS. If we find one, use
            // the hostname returned to construct an Autodiscover endpoint URL.
            autodiscoverUrl = GetRedirectionUrlFromDnsSrvRecord(domainName);
            if (autodiscoverUrl != null &&
                CallRedirectionUrlValidationCallback(autodiscoverUrl.ToString()) &&
                TryGetAutodiscoverEndpointUrl(autodiscoverUrl.Host, out autodiscoverUrl))
            {
                var result = await getSettingsMethod(identities, settings, requestedVersion, autodiscoverUrl);
                response = result.Item1;
                autodiscoverUrl = result.Item2;

                // If we got this far, the response was successful, set Url.
                Url = autodiscoverUrl;

                return response;
            }

            throw new AutodiscoverLocalException(Strings.AutodiscoverCouldNotBeLocated);
        }
    }

    /// <summary>
    ///     Gets settings for one or more users.
    /// </summary>
    /// <param name="smtpAddresses">The SMTP addresses of the users.</param>
    /// <param name="settings">The settings.</param>
    /// <param name="requestedVersion">Requested version of the Exchange service.</param>
    /// <param name="autodiscoverUrl">The autodiscover URL.</param>
    /// <returns>GetUserSettingsResponse collection.</returns>
    private async Task<Tuple<GetUserSettingsResponseCollection, Uri>> InternalGetUserSettings(
        List<string> smtpAddresses,
        List<UserSettingName> settings,
        ExchangeVersion? requestedVersion,
        Uri autodiscoverUrl
    )
    {
        // The response to GetUserSettings can be a redirection. Execute GetUserSettings until we get back 
        // a valid response or we've followed too many redirections.
        for (var currentHop = 0; currentHop < AutodiscoverMaxRedirections; currentHop++)
        {
            var request = new GetUserSettingsRequest(this, autodiscoverUrl)
            {
                SmtpAddresses = smtpAddresses,
                Settings = settings,
            };
            var response = await request.Execute();

            // Did we get redirected?
            if (response.ErrorCode == AutodiscoverErrorCode.RedirectUrl && response.RedirectionUrl != null)
            {
                TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    $"Request to {autodiscoverUrl} returned redirection to {response.RedirectionUrl}"
                );

                // this url need be brought back to the caller.
                //
                autodiscoverUrl = response.RedirectionUrl;
            }
            else
            {
                return Tuple.Create(response, autodiscoverUrl);
            }
        }

        TraceMessage(
            TraceFlags.AutodiscoverConfiguration,
            $"Maximum number of redirection hops {AutodiscoverMaxRedirections} exceeded"
        );

        throw new AutodiscoverLocalException(Strings.MaximumRedirectionHopsExceeded);
    }

    /// <summary>
    ///     Gets the domain settings using Autodiscover SOAP service.
    /// </summary>
    /// <param name="domains">The domains.</param>
    /// <param name="settings">The settings.</param>
    /// <param name="requestedVersion">Requested version of the Exchange service.</param>
    /// <returns>GetDomainSettingsResponse collection.</returns>
    internal Task<GetDomainSettingsResponseCollection> GetDomainSettings(
        List<string> domains,
        List<DomainSettingName> settings,
        ExchangeVersion? requestedVersion
    )
    {
        EwsUtilities.ValidateParam(domains);
        EwsUtilities.ValidateParam(settings);

        return GetSettings(domains, settings, requestedVersion, InternalGetDomainSettings, () => domains[0]);
    }

    /// <summary>
    ///     Gets settings for one or more domains.
    /// </summary>
    /// <param name="domains">The domains.</param>
    /// <param name="settings">The settings.</param>
    /// <param name="requestedVersion">Requested version of the Exchange service.</param>
    /// <param name="autodiscoverUrl">The autodiscover URL.</param>
    /// <returns>GetDomainSettingsResponse collection.</returns>
    private async Task<Tuple<GetDomainSettingsResponseCollection, Uri>> InternalGetDomainSettings(
        List<string> domains,
        List<DomainSettingName> settings,
        ExchangeVersion? requestedVersion,
        Uri autodiscoverUrl
    )
    {
        // The response to GetDomainSettings can be a redirection. Execute GetDomainSettings until we get back 
        // a valid response or we've followed too many redirections.
        for (var currentHop = 0; currentHop < AutodiscoverMaxRedirections; currentHop++)
        {
            var request = new GetDomainSettingsRequest(this, autodiscoverUrl)
            {
                Domains = domains,
                Settings = settings,
                RequestedVersion = requestedVersion,
            };
            var response = await request.Execute();

            // Did we get redirected?
            if (response.ErrorCode == AutodiscoverErrorCode.RedirectUrl && response.RedirectionUrl != null)
            {
                autodiscoverUrl = response.RedirectionUrl;
            }
            else
            {
                return Tuple.Create(response, autodiscoverUrl);
            }
        }

        TraceMessage(
            TraceFlags.AutodiscoverConfiguration,
            $"Maximum number of redirection hops {AutodiscoverMaxRedirections} exceeded"
        );

        throw new AutodiscoverLocalException(Strings.MaximumRedirectionHopsExceeded);
    }

    /// <summary>
    ///     Gets the autodiscover endpoint URL.
    /// </summary>
    /// <param name="host">The host.</param>
    /// <returns></returns>
    private Uri GetAutodiscoverEndpointUrl(string host)
    {
        if (TryGetAutodiscoverEndpointUrl(host, out var autodiscoverUrl))
        {
            return autodiscoverUrl;
        }

        throw new AutodiscoverLocalException(Strings.NoSoapOrWsSecurityEndpointAvailable);
    }

    /// <summary>
    ///     Tries the get Autodiscover Service endpoint URL.
    /// </summary>
    /// <param name="host">The host.</param>
    /// <param name="url">The URL.</param>
    /// <returns></returns>
    private bool TryGetAutodiscoverEndpointUrl(string host, [MaybeNullWhen(false)] out Uri url)
    {
        url = null;

        if (TryGetEnabledEndpointsForHost(ref host, out var endpoints))
        {
            url = new Uri(string.Format(AutodiscoverSoapHttpsUrl, host));

            // Make sure that at least one of the non-legacy endpoints is available.
            if ((endpoints & AutodiscoverEndpoints.Soap) != AutodiscoverEndpoints.Soap &&
                (endpoints & AutodiscoverEndpoints.WsSecurity) != AutodiscoverEndpoints.WsSecurity &&
                (endpoints & AutodiscoverEndpoints.WSSecuritySymmetricKey) !=
                AutodiscoverEndpoints.WSSecuritySymmetricKey &&
                (endpoints & AutodiscoverEndpoints.WSSecurityX509Cert) != AutodiscoverEndpoints.WSSecurityX509Cert &&
                (endpoints & AutodiscoverEndpoints.OAuth) != AutodiscoverEndpoints.OAuth)
            {
                TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    $"No Autodiscover endpoints are available  for host {host}"
                );

                return false;
            }

            // If we have WLID credentials, make sure that we have a WS-Security endpoint
            if (Credentials is WindowsLiveCredentials)
            {
                if ((endpoints & AutodiscoverEndpoints.WsSecurity) != AutodiscoverEndpoints.WsSecurity)
                {
                    TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        $"No Autodiscover WS-Security endpoint is available for host {host}"
                    );

                    return false;
                }

                url = new Uri(string.Format(AutodiscoverSoapWsSecurityHttpsUrl, host));
            }
            //todo: implement PartnerTokenCredentials and X509CertificateCredentials
            else if (Credentials is PartnerTokenCredentials)
            {
                if ((endpoints & AutodiscoverEndpoints.WSSecuritySymmetricKey) !=
                    AutodiscoverEndpoints.WSSecuritySymmetricKey)
                {
                    TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        $"No Autodiscover WS-Security/SymmetricKey endpoint is available for host {host}"
                    );

                    return false;
                }

                url = new Uri(string.Format(AutodiscoverSoapWsSecuritySymmetricKeyHttpsUrl, host));
            }
            else if (Credentials is X509CertificateCredentials)
            {
                if ((endpoints & AutodiscoverEndpoints.WSSecurityX509Cert) != AutodiscoverEndpoints.WSSecurityX509Cert)
                {
                    TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        $"No Autodiscover WS-Security/X509Cert endpoint is available for host {host}"
                    );

                    return false;
                }

                url = new Uri(string.Format(AutodiscoverSoapWsSecurityX509CertHttpsUrl, host));
            }
            else if (Credentials is OAuthCredentials)
            {
                // If the credential is OAuthCredentials, no matter whether we have
                // the corresponding x-header, we will go with OAuth. 
                url = new Uri(string.Format(AutodiscoverSoapHttpsUrl, host));
            }

            return true;
        }

        TraceMessage(TraceFlags.AutodiscoverConfiguration, $"No Autodiscover endpoints are available for host {host}");

        return false;
    }

    /// <summary>
    ///     Defaults the get autodiscover service urls for domain.
    /// </summary>
    /// <param name="domainName">Name of the domain.</param>
    /// <returns></returns>
    private ICollection<string> DefaultGetScpUrlsForDomain(string domainName)
    {
        var helper = new DirectoryHelper(this);
        return helper.GetAutodiscoverScpUrlsForDomain(domainName);
    }

    /// <summary>
    ///     Gets the list of autodiscover service URLs.
    /// </summary>
    /// <param name="domainName">Domain name.</param>
    /// <param name="scpHostCount">Count of hosts found via SCP lookup.</param>
    /// <returns>List of Autodiscover URLs.</returns>
    internal List<Uri> GetAutodiscoverServiceUrls(string domainName, out int scpHostCount)
    {
        var urls = new List<Uri>();

        if (EnableScpLookup)
        {
            // Get SCP URLs
            var callback = GetScpUrlsForDomainCallback ?? DefaultGetScpUrlsForDomain;
            var scpUrls = callback(domainName);
            foreach (var str in scpUrls)
            {
                urls.Add(new Uri(str));
            }
        }

        scpHostCount = urls.Count;

        // As a fallback, add autodiscover URLs base on the domain name.
        urls.Add(new Uri(string.Format(AutodiscoverLegacyHttpsUrl, domainName)));
        urls.Add(new Uri(string.Format(AutodiscoverLegacyHttpsUrl, "autodiscover." + domainName)));

        return urls;
    }

    /// <summary>
    ///     Gets the list of autodiscover service hosts.
    /// </summary>
    /// <param name="domainName">Name of the domain.</param>
    /// <param name="scpHostCount">Count of SCP hosts that were found.</param>
    /// <returns>List of host names.</returns>
    internal List<string> GetAutodiscoverServiceHosts(string domainName, out int scpHostCount)
    {
        var serviceHosts = new List<string>();
        foreach (var url in GetAutodiscoverServiceUrls(domainName, out scpHostCount))
        {
            serviceHosts.Add(url.Host);
        }

        return serviceHosts;
    }

    /// <summary>
    ///     Gets the enabled autodiscover endpoints on a specific host.
    /// </summary>
    /// <param name="host">The host.</param>
    /// <param name="endpoints">Endpoints found for host.</param>
    /// <returns>Flags indicating which endpoints are enabled.</returns>
    private bool TryGetEnabledEndpointsForHost(ref string host, out AutodiscoverEndpoints endpoints)
    {
        TraceMessage(TraceFlags.AutodiscoverConfiguration, $"Determining which endpoints are enabled for host {host}");

        // We may get redirected to another host. And therefore need to limit the number
        // of redirections we'll tolerate.
        for (var currentHop = 0; currentHop < AutodiscoverMaxRedirections; currentHop++)
        {
            var autoDiscoverUrl = new Uri(string.Format(AutodiscoverLegacyHttpsUrl, host));

            endpoints = AutodiscoverEndpoints.None;


            IEwsHttpWebResponse? response = null;

            try
            {
                using var client = new HttpClient(
                    new HttpClientHandler
                    {
                        AllowAutoRedirect = false,
                    }
                );
                client.Timeout = AutodiscoverTimeout;
                var httpResponse = client.GetAsync(autoDiscoverUrl).Result;
                response = new EwsHttpWebResponse(httpResponse);
            }
            catch (Exception ex)
            {
                TraceMessage(TraceFlags.AutodiscoverConfiguration, $"I/O error: {ex.Message}");
            }

            if (response != null)
            {
                using (response)
                {
                    if (TryGetRedirectionResponse(response, out var redirectUrl))
                    {
                        TraceMessage(
                            TraceFlags.AutodiscoverConfiguration,
                            $"Host returned redirection to host '{redirectUrl.Host}'"
                        );

                        host = redirectUrl.Host;
                    }
                    else
                    {
                        endpoints = GetEndpointsFromHttpWebResponse(response);

                        TraceMessage(
                            TraceFlags.AutodiscoverConfiguration,
                            $"Host returned enabled endpoint flags: {endpoints}"
                        );

                        return true;
                    }
                }
            }
            else
            {
                return false;
            }
        }

        TraceMessage(
            TraceFlags.AutodiscoverConfiguration,
            $"Maximum number of redirection hops {AutodiscoverMaxRedirections} exceeded"
        );

        throw new AutodiscoverLocalException(Strings.MaximumRedirectionHopsExceeded);
    }

    /// <summary>
    ///     Gets the endpoints from HTTP web response.
    /// </summary>
    /// <param name="response">The response.</param>
    /// <returns>Endpoints enabled.</returns>
    private static AutodiscoverEndpoints GetEndpointsFromHttpWebResponse(IEwsHttpWebResponse response)
    {
        var endpoints = AutodiscoverEndpoints.Legacy;
        if (response.Headers.TryGetValues(AutodiscoverSoapEnabledHeaderName, out _))
        {
            endpoints |= AutodiscoverEndpoints.Soap;
        }

        if (response.Headers.TryGetValues(AutodiscoverWsSecurityEnabledHeaderName, out _))
        {
            endpoints |= AutodiscoverEndpoints.WsSecurity;
        }

        if (response.Headers.TryGetValues(AutodiscoverWsSecuritySymmetricKeyEnabledHeaderName, out _))
        {
            endpoints |= AutodiscoverEndpoints.WSSecuritySymmetricKey;
        }

        if (response.Headers.TryGetValues(AutodiscoverWsSecurityX509CertEnabledHeaderName, out _))
        {
            endpoints |= AutodiscoverEndpoints.WSSecurityX509Cert;
        }

        if (response.Headers.TryGetValues(AutodiscoverOAuthEnabledHeaderName, out _))
        {
            endpoints |= AutodiscoverEndpoints.OAuth;
        }

        return endpoints;
    }

    /// <summary>
    ///     Traces the response.
    /// </summary>
    /// <param name="response">The response.</param>
    /// <param name="memoryStream">The response content in a MemoryStream.</param>
    internal void TraceResponse(IEwsHttpWebResponse response, MemoryStream memoryStream)
    {
        ProcessHttpResponseHeaders(TraceFlags.AutodiscoverResponseHttpHeaders, response);

        if (TraceEnabled)
        {
            if (!string.IsNullOrEmpty(response.ContentType) &&
                (response.ContentType.StartsWith("text/", StringComparison.OrdinalIgnoreCase) ||
                 response.ContentType.StartsWith("application/soap", StringComparison.OrdinalIgnoreCase)))
            {
                TraceXml(TraceFlags.AutodiscoverResponse, memoryStream);
            }
            else
            {
                TraceMessage(TraceFlags.AutodiscoverResponse, "Non-textual response");
            }
        }
    }

    #endregion


    #region Utilities

    /// <summary>
    ///     Creates an HttpWebRequest instance and initializes it with the appropriate parameters,
    ///     based on the configuration of this service object.
    /// </summary>
    /// <param name="url">The URL that the HttpWebRequest should target.</param>
    /// <returns>A initialized instance of HttpWebRequest.</returns>
    internal HttpRequestMessage PrepareHttpRequestMessageForUrl(Uri url)
    {
        // Verify that the protocol is something that we can handle
        if (url.Scheme != "http" && url.Scheme != "https")
        {
            throw new ServiceLocalException(string.Format(Strings.UnsupportedWebProtocol, url.Scheme));
        }

        var request = new HttpRequestMessage(HttpMethod.Post, url);

        request.Headers.Accept.ParseAdd("text/xml");
        request.Headers.UserAgent.ParseAdd(UserAgent);

        if (!string.IsNullOrEmpty(ClientRequestId))
        {
            request.Headers.Add("client-request-id", ClientRequestId);
            if (ReturnClientRequestId)
            {
                request.Headers.Add("return-client-request-id", "true");
            }
        }

        if (HttpHeaders.Count > 0)
        {
            HttpHeaders.ForEach(kv => request.Headers.Add(kv.Key, kv.Value));
        }

        HttpResponseHeaders.Clear();

        return request;
    }

    internal HttpClient PrepareHttpClient()
    {
        var httpClientHandler = new HttpClientHandler
        {
            PreAuthenticate = PreAuthenticate,
            AllowAutoRedirect = false,
            CookieContainer = CookieContainer,
            UseDefaultCredentials = UseDefaultCredentials,
        };

        if (WebProxy != null)
        {
            httpClientHandler.Proxy = WebProxy;
            httpClientHandler.UseProxy = true;
        }

        if (!UseDefaultCredentials)
        {
            var serviceCredentials = Credentials;
            if (serviceCredentials == null)
            {
                throw new ServiceLocalException(Strings.CredentialsRequired);
            }

            // Temporary fix for authentication on Linux platform
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                serviceCredentials = AdjustLinuxAuthentication(_url, serviceCredentials);
            }

            // Make sure that credentials have been authenticated if required
            serviceCredentials.PreAuthenticate();

            // TODO support different credentials
            if (serviceCredentials is not WebCredentials)
            {
                throw new NotImplementedException();
            }

            httpClientHandler.Credentials = (Credentials as WebCredentials)?.Credentials;

            // Apply credentials to the request
            // serviceCredentials.PrepareWebRequest(request);
        }

        return new HttpClient(httpClientHandler)
        {
            Timeout = TimeSpan.FromMilliseconds(Timeout),
        };
    }

    /// <summary>
    ///     Calls the redirection URL validation callback.
    /// </summary>
    /// <param name="redirectionUrl">The redirection URL.</param>
    /// <remarks>
    ///     If the redirection URL validation callback is null, use the default callback which
    ///     does not allow following any redirections.
    /// </remarks>
    /// <returns>True if redirection should be followed.</returns>
    private bool CallRedirectionUrlValidationCallback(string redirectionUrl)
    {
        var callback = RedirectionUrlValidationCallback == null ? DefaultAutodiscoverRedirectionUrlValidationCallback
            : RedirectionUrlValidationCallback;
        return callback(redirectionUrl);
    }

    /// <summary>
    ///     Processes an HTTP error response.
    /// </summary>
    /// <param name="httpWebResponse">The HTTP web response.</param>
    /// <param name="webException">The web exception.</param>
    internal override void ProcessHttpErrorResponse(
        IEwsHttpWebResponse httpWebResponse,
        EwsHttpClientException webException
    )
    {
        InternalProcessHttpErrorResponse(
            httpWebResponse,
            webException,
            TraceFlags.AutodiscoverResponseHttpHeaders,
            TraceFlags.AutodiscoverResponse
        );
    }

    #endregion


    #region Constructors

    /// <summary>
    ///     Initializes a new instance of the <see cref="AutodiscoverService" /> class.
    /// </summary>
    public AutodiscoverService()
        : this(ExchangeVersion.Exchange2010)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AutodiscoverService" /> class.
    /// </summary>
    /// <param name="requestedServerVersion">The requested server version.</param>
    public AutodiscoverService(ExchangeVersion requestedServerVersion)
        : this(null, null, requestedServerVersion)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AutodiscoverService" /> class.
    /// </summary>
    /// <param name="domain">The domain that will be used to determine the URL of the service.</param>
    public AutodiscoverService(string domain)
        : this(null, domain)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AutodiscoverService" /> class.
    /// </summary>
    /// <param name="domain">The domain that will be used to determine the URL of the service.</param>
    /// <param name="requestedServerVersion">The requested server version.</param>
    public AutodiscoverService(string domain, ExchangeVersion requestedServerVersion)
        : this(null, domain, requestedServerVersion)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AutodiscoverService" /> class.
    /// </summary>
    /// <param name="url">The URL of the service.</param>
    public AutodiscoverService(Uri url)
        : this(url, url.Host)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AutodiscoverService" /> class.
    /// </summary>
    /// <param name="url">The URL of the service.</param>
    /// <param name="requestedServerVersion">The requested server version.</param>
    public AutodiscoverService(Uri url, ExchangeVersion requestedServerVersion)
        : this(url, url.Host, requestedServerVersion)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AutodiscoverService" /> class.
    /// </summary>
    /// <param name="url">The URL of the service.</param>
    /// <param name="domain">The domain that will be used to determine the URL of the service.</param>
    internal AutodiscoverService(Uri url, string domain)
    {
        EwsUtilities.ValidateDomainNameAllowNull(domain);

        _url = url;
        _domain = domain;
        _dnsClient = new AutodiscoverDnsClient(this);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AutodiscoverService" /> class.
    /// </summary>
    /// <param name="url">The URL of the service.</param>
    /// <param name="domain">The domain that will be used to determine the URL of the service.</param>
    /// <param name="requestedServerVersion">The requested server version.</param>
    internal AutodiscoverService(Uri url, string domain, ExchangeVersion requestedServerVersion)
        : base(requestedServerVersion)
    {
        EwsUtilities.ValidateDomainNameAllowNull(domain);

        _url = url;
        _domain = domain;
        _dnsClient = new AutodiscoverDnsClient(this);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AutodiscoverService" /> class.
    /// </summary>
    /// <param name="service">The other service.</param>
    /// <param name="requestedServerVersion">The requested server version.</param>
    internal AutodiscoverService(ExchangeServiceBase service, ExchangeVersion requestedServerVersion)
        : base(service, requestedServerVersion)
    {
        _dnsClient = new AutodiscoverDnsClient(this);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AutodiscoverService" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    internal AutodiscoverService(ExchangeServiceBase service)
        : this(service, service.RequestedServerVersion)
    {
    }

    #endregion


    #region Public Methods

    /// <summary>
    ///     Retrieves the specified settings for single SMTP address.
    /// </summary>
    /// <param name="userSmtpAddress">The SMTP addresses of the user.</param>
    /// <param name="userSettingNames">The user setting names.</param>
    /// <returns>A UserResponse object containing the requested settings for the specified user.</returns>
    /// <remarks>
    ///     This method handles will run the entire Autodiscover "discovery" algorithm and will follow address and URL
    ///     redirections.
    /// </remarks>
    public async Task<GetUserSettingsResponse> GetUserSettings(
        string userSmtpAddress,
        params UserSettingName[] userSettingNames
    )
    {
        var requestedSettings = new List<UserSettingName>(userSettingNames);

        if (string.IsNullOrEmpty(userSmtpAddress))
        {
            throw new ServiceValidationException(Strings.InvalidAutodiscoverSmtpAddress);
        }

        if (requestedSettings.Count == 0)
        {
            throw new ServiceValidationException(Strings.InvalidAutodiscoverSettingsCount);
        }

        if (RequestedServerVersion < MinimumRequestVersionForAutoDiscoverSoapService)
        {
            return await InternalGetLegacyUserSettings(userSmtpAddress, requestedSettings);
        }

        return await InternalGetSoapUserSettings(userSmtpAddress, requestedSettings);
    }

    /// <summary>
    ///     Retrieves the specified settings for a set of users.
    /// </summary>
    /// <param name="userSmtpAddresses">The SMTP addresses of the users.</param>
    /// <param name="userSettingNames">The user setting names.</param>
    /// <returns>A GetUserSettingsResponseCollection object containing the responses for each individual user.</returns>
    public Task<GetUserSettingsResponseCollection> GetUsersSettings(
        IEnumerable<string> userSmtpAddresses,
        params UserSettingName[] userSettingNames
    )
    {
        if (RequestedServerVersion < MinimumRequestVersionForAutoDiscoverSoapService)
        {
            throw new ServiceVersionException(
                string.Format(
                    Strings.AutodiscoverServiceIncompatibleWithRequestVersion,
                    MinimumRequestVersionForAutoDiscoverSoapService
                )
            );
        }

        var smtpAddresses = new List<string>(userSmtpAddresses);
        var settings = new List<UserSettingName>(userSettingNames);

        return GetUserSettings(smtpAddresses, settings);
    }

    /// <summary>
    ///     Retrieves the specified settings for a domain.
    /// </summary>
    /// <param name="domain">The domain.</param>
    /// <param name="requestedVersion">Requested version of the Exchange service.</param>
    /// <param name="domainSettingNames">The domain setting names.</param>
    /// <returns>A DomainResponse object containing the requested settings for the specified domain.</returns>
    public async Task<GetDomainSettingsResponse> GetDomainSettings(
        string domain,
        ExchangeVersion? requestedVersion,
        params DomainSettingName[] domainSettingNames
    )
    {
        var domains = new List<string>(1)
        {
            domain,
        };
        var settings = new List<DomainSettingName>(domainSettingNames);
        return (await GetDomainSettings(domains, settings, requestedVersion))[0];
    }

    /// <summary>
    ///     Retrieves the specified settings for a set of domains.
    /// </summary>
    /// <param name="domains">The SMTP addresses of the domains.</param>
    /// <param name="requestedVersion">Requested version of the Exchange service.</param>
    /// <param name="domainSettingNames">The domain setting names.</param>
    /// <returns>A GetDomainSettingsResponseCollection object containing the responses for each individual domain.</returns>
    public Task<GetDomainSettingsResponseCollection> GetDomainSettings(
        IEnumerable<string> domains,
        ExchangeVersion? requestedVersion,
        params DomainSettingName[] domainSettingNames
    )
    {
        var settings = new List<DomainSettingName>(domainSettingNames);

        return GetDomainSettings(new List<string>(domains), settings, requestedVersion);
    }

    /// <summary>
    ///     Try to get the partner access information for the given target tenant.
    /// </summary>
    /// <param name="targetTenantDomain">The target domain or user email address.</param>
    /// <param name="partnerAccessCredentials">The partner access credentials.</param>
    /// <param name="targetTenantAutodiscoverUrl">The autodiscover url for the given tenant.</param>
    /// <returns>True if the partner access information was retrieved, false otherwise.</returns>
    public async Task<Tuple<bool, ExchangeCredentials, Uri>> TryGetPartnerAccess(string targetTenantDomain)
    {
        EwsUtilities.ValidateNonBlankStringParam(targetTenantDomain, nameof(targetTenantDomain));

        // the user should set the url to its own tenant's autodiscover url.
        // 
        if (Url == null)
        {
            throw new ServiceValidationException(Strings.PartnerTokenRequestRequiresUrl);
        }

        if (RequestedServerVersion < ExchangeVersion.Exchange2010_SP1)
        {
            throw new ServiceVersionException(
                string.Format(Strings.PartnerTokenIncompatibleWithRequestVersion, ExchangeVersion.Exchange2010_SP1)
            );
        }

        ExchangeCredentials? partnerAccessCredentials = null;
        Uri? targetTenantAutodiscoverUrl = null;

        var smtpAddress = targetTenantDomain;
        if (!smtpAddress.Contains('@'))
        {
            smtpAddress = "SystemMailbox{e0dc1c29-89c3-4034-b678-e6c29d823ed9}@" + targetTenantDomain;
        }

        var request = new GetUserSettingsRequest(this, Url, true)
        {
            SmtpAddresses = new List<string>(
                new[]
                {
                    smtpAddress,
                }
            ),
            Settings = new List<UserSettingName>(
                new[]
                {
                    UserSettingName.ExternalEwsUrl,
                }
            ),
        };

        GetUserSettingsResponseCollection? response;
        try
        {
            response = await request.Execute();
        }
        catch (ServiceRequestException)
        {
            return Tuple.Create(false, partnerAccessCredentials, targetTenantAutodiscoverUrl);
        }
        catch (ServiceRemoteException)
        {
            return Tuple.Create(false, partnerAccessCredentials, targetTenantAutodiscoverUrl);
        }

        if (string.IsNullOrEmpty(request.PartnerToken) || string.IsNullOrEmpty(request.PartnerTokenReference))
        {
            return Tuple.Create(false, partnerAccessCredentials, targetTenantAutodiscoverUrl);
        }

        if (response.ErrorCode == AutodiscoverErrorCode.NoError)
        {
            var firstResponse = response.Responses[0];
            if (firstResponse.ErrorCode == AutodiscoverErrorCode.NoError)
            {
                targetTenantAutodiscoverUrl = Url;
            }
            else if (firstResponse.ErrorCode == AutodiscoverErrorCode.RedirectUrl)
            {
                targetTenantAutodiscoverUrl = new Uri(firstResponse.RedirectTarget);
            }
            else
            {
                return Tuple.Create(false, partnerAccessCredentials, targetTenantAutodiscoverUrl);
            }
        }
        else
        {
            return Tuple.Create(false, partnerAccessCredentials, targetTenantAutodiscoverUrl);
        }

        partnerAccessCredentials = new PartnerTokenCredentials(request.PartnerToken, request.PartnerTokenReference);

        targetTenantAutodiscoverUrl = partnerAccessCredentials.AdjustUrl(targetTenantAutodiscoverUrl);

        return Tuple.Create(true, partnerAccessCredentials, targetTenantAutodiscoverUrl);
    }

    #endregion


    #region Properties

    /// <summary>
    ///     Gets or sets the domain this service is bound to. When this property is set, the domain
    ///     name is used to automatically determine the Autodiscover service URL.
    /// </summary>
    public string? Domain
    {
        get => _domain;
        set
        {
            EwsUtilities.ValidateDomainNameAllowNull(value, "Domain");

            // If Domain property is set to non-null value, Url property is nulled.
            if (value != null)
            {
                _url = null;
            }

            _domain = value;
        }
    }

    /// <summary>
    ///     Gets or sets the URL this service is bound to.
    /// </summary>
    public Uri? Url
    {
        get => _url;
        set
        {
            // If Url property is set to non-null value, Domain property is set to host portion of Url.
            if (value != null)
            {
                _domain = value.Host;
            }

            _url = value;
        }
    }

    /// <summary>
    ///     Gets a value indicating whether the Autodiscover service that URL points to is internal (inside the corporate
    ///     network)
    ///     or external (outside the corporate network).
    /// </summary>
    /// <remarks>
    ///     IsExternal is null in the following cases:
    ///     - This instance has been created with a domain name and no method has been called,
    ///     - This instance has been created with a URL.
    /// </remarks>
    public bool? IsExternal { get; internal set; } = true;

    /// <summary>
    ///     Gets or sets the redirection URL validation callback.
    /// </summary>
    /// <value>The redirection URL validation callback.</value>
    public AutodiscoverRedirectionUrlValidationCallback RedirectionUrlValidationCallback { get; set; }

    /// <summary>
    ///     Gets or sets the DNS server address.
    /// </summary>
    /// <value>The DNS server address.</value>
    internal IPAddress DnsServerAddress { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether the AutodiscoverService should perform SCP (ServiceConnectionPoint) record
    ///     lookup when determining
    ///     the Autodiscover service URL.
    /// </summary>
    public bool EnableScpLookup { get; set; } = true;

    /// <summary>
    ///     Gets or sets the delegate used to resolve Autodiscover SCP urls for a specified domain.
    /// </summary>
    public Func<string, ICollection<string>> GetScpUrlsForDomainCallback { get; set; }

    #endregion
}
