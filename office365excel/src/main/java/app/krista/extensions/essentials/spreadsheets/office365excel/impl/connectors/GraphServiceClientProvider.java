/*
 * Office365excel Extension for Krista
 * Copyright (C) 2025 Krista Software
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

package app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors;

import app.krista.extension.authorization.MustAuthorizeException;
import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.RefreshTokenStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.AuthErrorRule;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.ksdk.context.AuthorizationContext;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.field.NamedValuedField;
import com.google.gson.JsonObject;
import com.microsoft.aad.msal4j.*;
import com.microsoft.graph.authentication.IAuthenticationProvider;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.httpcore.HttpClients;
import com.microsoft.graph.models.User;
import com.microsoft.graph.requests.DriveRequestBuilder;
import com.microsoft.graph.requests.GraphServiceClient;
import com.microsoft.graph.requests.WorkbookWorksheetRequestBuilder;
import okhttp3.OkHttpClient;
import org.jetbrains.annotations.NotNull;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.IOException;
import java.net.URL;
import java.util.*;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeUnit;
import java.util.stream.Collectors;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.MicrosoftAuthErrorConstants.AUTH_ERROR_RULES;
import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.*;

public class GraphServiceClientProvider {

    private final RefreshTokenStore refreshTokenStore;
    private final OfficeExcelAttributes officeExcelAttributes;
    private final RequestContext requestContext;
    private final AuthorizationContext authorizationContext;
    private final String authContextId;
    private static final Logger LOGGER = LoggerFactory.getLogger(GraphServiceClientProvider.class);

    @Inject
    public GraphServiceClientProvider(RefreshTokenStore refreshTokenStore, OfficeExcelAttributes officeExcelAttributes,
                                      RequestContext requestContext, AuthorizationContext authorizationContext) {
        this(refreshTokenStore, officeExcelAttributes, requestContext, authorizationContext, null);
    }

    public GraphServiceClientProvider(RefreshTokenStore refreshTokenStore, OfficeExcelAttributes officeExcelAttributes,
                                      RequestContext requestContext, AuthorizationContext authorizationContext, String authContextId) {
        this.refreshTokenStore = refreshTokenStore;
        this.officeExcelAttributes = officeExcelAttributes;
        this.requestContext = requestContext;
        this.authorizationContext = authorizationContext;
        this.authContextId = authContextId;
    }

    public OfficeExcelAttributes getOfficeExcelAttributes() {
        return officeExcelAttributes;
    }

    /**
     * Retrieves a Graph Service Client for asynchronous operations.
     *
     * @param useEmail true to use admin email for authentication, false for user context
     * @param accountId the account ID for authentication (null to use default)
     * @return GraphServiceClient configured with appropriate authentication
     * @throws IOException if authentication fails or refresh token is not found
     */
    public GraphServiceClient getGraphServiceClientForAsync(boolean useEmail, String accountId) throws IOException {
        return getGraphServiceClient(useEmail, accountId);
    }

    /**
     * Retrieves a Graph Service Client using the provided request context.
     * Determines authentication flow based on the invokeAsUser flag in the context.
     *
     * @param threadLocalRequestContext the request context containing authentication information
     * @return GraphServiceClient configured with appropriate authentication
     * @throws IOException if authentication fails or refresh token is not found
     */
    public GraphServiceClient getGraphServiceClient(RequestContext threadLocalRequestContext) throws IOException {
        return getGraphServiceClient(!threadLocalRequestContext.invokeAsUser(), null);
    }

    /**
     * Retrieves a Graph Service Client using the default request context.
     * Determines authentication flow based on the invokeAsUser flag.
     *
     * @return GraphServiceClient configured with appropriate authentication
     * @throws IOException if authentication fails or refresh token is not found
     */
    public GraphServiceClient getGraphServiceClient() throws IOException {
        return getGraphServiceClient(!requestContext.invokeAsUser(), null);
    }

    /**
     * Retrieves a Graph Service Client configured for admin (service account) authentication.
     * Uses the admin email for authentication instead of user context.
     *
     * @return GraphServiceClient configured with admin authentication
     * @throws IOException if authentication fails or refresh token is not found
     */
    public GraphServiceClient getGraphServiceClientForAdmin() throws IOException {
        return getGraphServiceClient(true, null);
    }

    /**
     * Internal method to retrieve a Graph Service Client with specified authentication parameters.
     * Retrieves the refresh token from the store and creates an authenticated client.
     *
     * @param useEmail true to use admin email for authentication, false for user context
     * @param accountId the account ID for authentication (null to use default)
     * @return GraphServiceClient configured with appropriate authentication
     * @throws IOException if authentication fails or refresh token is not found
     * @throws MustAuthorizeException if refresh token is not found in the store
     */
    private GraphServiceClient getGraphServiceClient(boolean useEmail, String accountId) throws IOException {
        try {
            String userId = getUserId(useEmail, accountId);
            String refreshToken = refreshTokenStore.get(userId);
            if (refreshToken == null) {
                throw createMustAuthorizationException(userId, false);
            }
            return getGraphServiceClient(userId, refreshToken);
        } catch (IOException | RuntimeException cause) {
            LOGGER.error("Failed to get graph client {}", cause.getClass().getName());
            throw cause;
        }
    }

    /**
     * Creates a Graph Service Client using a refresh token.
     * Acquires an access token using MSAL4J and configures the client with custom HTTP settings.
     * Automatically stores updated refresh tokens for token rotation.
     *
     * @param refreshTokenStoreKey the key used to store the refresh token
     * @param refreshToken the OAuth 2.0 refresh token
     * @return GraphServiceClient configured with the acquired access token
     * @throws RuntimeException if token acquisition fails
     */
    private GraphServiceClient getGraphServiceClient(String refreshTokenStoreKey, String refreshToken) {
        try {
            String[] scopes = Constants.SCOPE.split(Constants.SPACE);
            Set<String> scopeSet = Arrays.stream(scopes).collect(Collectors.toSet());
            RefreshTokenParameters refreshTokenParameters = RefreshTokenParameters.builder(scopeSet, refreshToken).build();
            IClientCredential clientCredential = ClientCredentialFactory.createFromSecret(officeExcelAttributes.getClientSecret());
            String authority = Constants.AUTHORITY + officeExcelAttributes.getTenantId();

            ConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplication.builder
                    (officeExcelAttributes.getClientId(), clientCredential).authority(authority).build();
            IAuthenticationResult authenticationResult = confidentialClientApplication.acquireToken(refreshTokenParameters).get();
            final String cachedTokenContent = confidentialClientApplication.tokenCache().serialize();
            updateRefreshToken(refreshTokenStoreKey, Constants.GSON.fromJson(cachedTokenContent, JsonObject.class));
            GraphServiceClientAuthenticationProvider authenticationProvider = new GraphServiceClientAuthenticationProvider(authenticationResult.accessToken());
            OkHttpClient customHttpClient = HttpClients.createDefault(authenticationProvider)
                    .newBuilder()
                    .connectTimeout(60, TimeUnit.SECONDS)  // default 10s → increase
                    .readTimeout(120, TimeUnit.SECONDS)   // increase for large responses
                    .writeTimeout(120, TimeUnit.SECONDS)  // increase for uploads
                    .build();

            return GraphServiceClient.builder()
                    .httpClient(customHttpClient)
                    .authenticationProvider(authenticationProvider)
                    .buildClient();
        } catch (ClientException | ExecutionException | IOException cause) {
            handleAuthenticationError(cause, refreshTokenStoreKey);
            // This line will never be reached as handleAuthenticationError always throws an exception
            return null;
        } catch (InterruptedException cause) {
            Thread.currentThread().interrupt();
            throw createMustAuthorizationException(refreshTokenStoreKey, true);
        }
    }

    private void updateRefreshToken(String userId, JsonObject jsonTokenCache) throws IOException {
        JsonObject refreshToken = jsonTokenCache.get("RefreshToken").getAsJsonObject();
        String firstKey = refreshToken.keySet().iterator().next();
        JsonObject tokenData = refreshToken.getAsJsonObject(firstKey);
        String secret = tokenData.get("secret").getAsString();
        refreshTokenStore.put(userId, secret);
        LOGGER.info("Updated Refresh Token.");
    }

    private MustAuthorizeException createMustAuthorizationException(String userId, boolean reAuthentication) {
        LOGGER.error(reAuthentication
                ? "Got error for Authentication. So sending for re authentication"
                : "Got error for Authentication. So sending for authentication");

        List<NamedValuedField> details = new ArrayList<>();
        NamedValuedField userIdField = new NamedValuedField("userId", "Text", userId, new HashMap<>(),
                new HashMap<>());
        details.add(userIdField);
        if (authContextId != null) {
            NamedValuedField contextIdField = new NamedValuedField("authContextId", "Text",
                    authContextId, new HashMap<>(), new HashMap<>());
            details.add(contextIdField);
        }
        return new MustAuthorizeException(reAuthentication
                ? Constants.REFRESH_TOKEN_EXPIRED_PLEASE_REAUTHORIZE_YOURSELF
                : Constants.AUTHORIZATION_PROMPT, details);
    }

    private String getUserId(boolean calledFromValidateAttributes, String accountId) throws IOException {
        String userId = calledFromValidateAttributes ? officeExcelAttributes.getEmailId()
                : (accountId == null ? authorizationContext.getAuthorizedAccount().getAccountId() : accountId);
        if (userId == null) {
            throw new IllegalStateException("Failed to get account");
        }
        return userId;
    }

    public static class GraphServiceClientAuthenticationProvider implements IAuthenticationProvider {

        private final String accessToken;

        public GraphServiceClientAuthenticationProvider(String accessToken) {
            if (accessToken == null || accessToken.isBlank()) {
                throw new IllegalArgumentException("Empty access token provided " + accessToken);
            }
            this.accessToken = accessToken;
        }

        @NotNull
        @Override
        public CompletableFuture<String> getAuthorizationTokenAsync(@NotNull URL requestUrl) {
            return CompletableFuture.completedFuture(accessToken);
        }
    }

    /**
     * Retrieves the email address of the authenticated admin user.
     * Uses the admin Graph Service Client to query the /me endpoint.
     *
     * @return the email address of the authenticated user, empty string if not found
     * @throws IOException if authentication fails or the API call fails
     */
    public String getAuthenticatedUserEmail() throws IOException {
        User user = getGraphServiceClientForAdmin().me().buildRequest().get();
        return user != null ? user.mail : "";
    }

    /**
     * Creates a DriveRequestBuilder for accessing OneDrive resources.
     * Supports both default drive and specific drive access with custom authentication.
     *
     * @param driveId the ID of the specific drive (null for default drive)
     * @param accountId the account ID for authentication (null to use default)
     * @param invokeAsUser true to invoke as user context, false for admin context
     * @return DriveRequestBuilder configured for the specified drive
     * @throws IOException if authentication fails
     */
    public DriveRequestBuilder getDriveRequestBuilder(String driveId, String accountId, Boolean invokeAsUser) throws IOException {
        return StringUtil.isNullOrEmpty(driveId) ? getGraphServiceClient().me().drive() :
                getGraphServiceClientForAsync(invokeAsUser, accountId).me().drives(driveId);
    }

    /**
     * Creates a DriveRequestBuilder for accessing OneDrive resources.
     * Uses default authentication context.
     *
     * @param driveId the ID of the specific drive (null for default drive)
     * @return DriveRequestBuilder configured for the specified drive
     * @throws RuntimeException if authentication fails
     */
    public DriveRequestBuilder getDriveRequestBuilder(String driveId) {
        try {
            return StringUtil.isNullOrEmpty(driveId) ? getGraphServiceClient().me().drive() : getGraphServiceClient().me().drives(driveId);
        } catch (IOException cause) {
            throw new RuntimeException(cause);
        }
    }

    public DriveRequestBuilder getDriveRequestBuilder(String driveId, RequestContext requestContext) {
        try {
            return StringUtil.isNullOrEmpty(driveId) ? getGraphServiceClient(requestContext).me().drive() : getGraphServiceClient(requestContext).me().drives(driveId);
        } catch (IOException cause) {
            throw new RuntimeException(cause);
        }
    }

    public WorkbookWorksheetRequestBuilder worksheetRequestBuilders(String driveId, String worksheetName, String workbookId) {
        return getDriveRequestBuilder(driveId)
                .items(workbookId)
                .workbook()
                .worksheets(worksheetName);
    }

    /**
     * Handles authentication errors and throws appropriate MustAuthorizeException
     *
     * @param cause                The exception that occurred during authentication
     * @param refreshTokenStoreKey The key used to store the refresh token
     * @throws MustAuthorizeException with appropriate error message based on the cause
     */
    private void handleAuthenticationError(Exception cause, String refreshTokenStoreKey) {
        String errorMessage = cause.getMessage();
        LOGGER.error(" handleAuthenticationError() -> errorMessage ::: {} ", errorMessage);
        if (errorMessage != null) {
            for (AuthErrorRule rule : AUTH_ERROR_RULES) {
                if (rule.matches(errorMessage)) {
                    if (rule.shouldRemoveToken()) {
                        refreshTokenStore.remove(refreshTokenStoreKey);
                    }
                    throw new MustAuthorizeException(rule.getUserMessage(), createAuthDetails(refreshTokenStoreKey));
                }
            }
        }
        throw createMustAuthorizationException(refreshTokenStoreKey, !requestContext.invokeAsUser());
    }

    private List<NamedValuedField> createAuthDetails(String userId) {
        List<NamedValuedField> details = new ArrayList<>();
        NamedValuedField userIdField = new NamedValuedField(USER_ID, TEXT, userId, new HashMap<>(), new HashMap<>());
        details.add(userIdField);
        if (authContextId != null) {
            NamedValuedField contextIdField = new NamedValuedField(AUTH_CONTEXT_ID, TEXT, authContextId, new HashMap<>(), new HashMap<>());
            details.add(contextIdField);
        }
        return details;
    }

}
