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

package app.krista.extensions.essentials.spreadsheets.office365excel.api;

import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.RefreshTokenStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.ksdk.authentication.AuthorizationListener;
import app.krista.ksdk.context.AuthorizationContext;
import com.github.scribejava.core.model.OAuth2AccessToken;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import javax.ws.rs.BadRequestException;
import javax.ws.rs.GET;
import javax.ws.rs.Path;
import javax.ws.rs.QueryParam;
import java.util.Objects;

@Path("/")
public final class AuthenticationResource {
    private final OfficeExcelAttributeStore officeExcelAttributeStore;
    private final RefreshTokenStore refreshTokenStore;
    private final GraphServiceClientProviderFactory providerFactory;
    private static final Logger LOGGER = LoggerFactory.getLogger(AuthenticationResource.class);

    private final AuthorizationContext context;

    private final AuthorizationListener authorizationListener;

    @Inject
    public AuthenticationResource(OfficeExcelAttributeStore officeExcelAttributeStore, RefreshTokenStore refreshTokenStore,
                                  GraphServiceClientProviderFactory providerFactory, AuthorizationContext context, AuthorizationListener authorizationListener) {
        this.officeExcelAttributeStore = officeExcelAttributeStore;
        this.refreshTokenStore = refreshTokenStore;
        this.context = context;
        this.authorizationListener = authorizationListener;
        this.providerFactory = providerFactory;
    }

    @GET
    @Path("callback")
    public String getCallBack(@QueryParam(Constants.CODE) String code, @QueryParam(Constants.STATE) String state) {
        Objects.requireNonNull(code);
        String[] parts = state.split("#");
        if (parts[0].isBlank() || parts.length > 2) {
            throw new BadRequestException("Authentication failed: Invalid authorization state received from Microsoft. " +
                    "Please try authenticating again. If the problem persists, contact your administrator.");
        }
        String key = parts[0];
        String authContextId = parts.length == 2 ? parts[1] : null;
        try {

            GraphServiceClientProvider clientProvider;
            if (authContextId == null) {
                clientProvider = providerFactory.create();
            } else {
                OfficeExcelAttributes effectiveOfficeExcelAttributes = officeExcelAttributeStore.load(authContextId);
                clientProvider = providerFactory.create(effectiveOfficeExcelAttributes);
            }
            OAuth2AccessToken accessToken = clientProvider.getOfficeExcelAttributes().getOAuth20Service().getAccessToken(code);
            refreshTokenStore.put(key, accessToken.getRefreshToken());
            if (!key.startsWith("wsContact") && !key.equalsIgnoreCase(clientProvider.getAuthenticatedUserEmail())) {
                refreshTokenStore.remove(key);
                return "Authentication failed: The Microsoft account you signed in with does not match the email configured in the extension setup. " +
                        "Expected: '" + key + "', but you signed in with a different account. " +
                        "Please sign in with the correct Microsoft account or update the extension configuration.";
            }
            if (context.isAuthenticated()) {
                authorizationListener.authorized();
                return "User authenticated successfully.";
            }

            return "User authenticated successfully. Please save the changes.";

        } catch (InterruptedException cause) {
            LOGGER.error("Authentication interrupted: {}", cause.getMessage(), cause);
            Thread.currentThread().interrupt();
            throw new IllegalStateException("Authentication was interrupted. Please try again. " +
                    "If this problem continues, contact your administrator.", cause.getCause());
        } catch (Exception exception) {
            LOGGER.error("Authentication failed: {}", exception.getMessage(), exception);
            throw new IllegalStateException("Authentication failed due to an unexpected error. " +
                    "Please verify your Microsoft account credentials and try again. " +
                    "If the problem persists, contact your administrator with this error: " + exception.getMessage(),
                    exception.getCause());
        } finally {
            if (authContextId != null) {
                officeExcelAttributeStore.remove(authContextId);
            }
        }

    }

}