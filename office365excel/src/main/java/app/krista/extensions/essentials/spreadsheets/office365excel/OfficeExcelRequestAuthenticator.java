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

package app.krista.extensions.essentials.spreadsheets.office365excel;

import app.krista.extension.authorization.MustAuthenticateException;
import app.krista.extension.authorization.MustAuthorizeException;
import app.krista.extension.authorization.RequestAuthenticator;
import app.krista.extension.request.ProtoRequest;
import app.krista.extension.request.ProtoResponse;
import app.krista.extension.request.protos.http.HttpRequest;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.model.field.NamedField;
import app.krista.model.field.NamedValuedField;
import com.github.scribejava.core.oauth.OAuth20Service;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import javax.ws.rs.core.MultivaluedMap;
import java.io.IOException;
import java.util.*;

import static com.github.scribejava.core.model.OAuthConstants.STATE;

@Service
public class OfficeExcelRequestAuthenticator implements RequestAuthenticator {
    private final OfficeExcelAttributes officeExcelAttributes;
    private final OfficeExcelAttributeStore officeExcelAttributeStore;
    private static final Logger LOGGER = LoggerFactory.getLogger(OfficeExcelRequestAuthenticator.class);


    @Inject
    public OfficeExcelRequestAuthenticator(OfficeExcelAttributes officeExcelAttributes, OfficeExcelAttributeStore officeExcelAttributeStore) {
        this.officeExcelAttributes = officeExcelAttributes;
        this.officeExcelAttributeStore = officeExcelAttributeStore;
    }

    @Override
    public String getScheme() {
        return null;
    }

    @Override
    public Set<String> getSupportedProtocols() {
        return null;
    }

    @Override
    public String getAuthenticatedAccountId(ProtoRequest protoRequest) {
        try {
            HttpRequest httpRequest = (HttpRequest) protoRequest;
            LOGGER.info("path correction {}", httpRequest.getUri().getPath());
            if (("/callback").equals(httpRequest.getUri().getPath())) {
                httpRequest.bufferBody();
                MultivaluedMap<String, String> queryParameters = httpRequest.getQueryParameters();
                String state = String.valueOf(queryParameters.get(STATE).getFirst());
                String[] parts = state.split("#");
                return parts[0];
            } else {
                return null;
            }
        } catch (RuntimeException | IOException cause) {
            return null;
        }
    }

    @Override
    public boolean setServiceAuthorization(String s) {
        return false;
    }

    @Override
    public Map<String, NamedField> getAttributeFields() {
        return null;
    }

    @Override
    public ProtoResponse getMustAuthenticateResponse(MustAuthenticateException cause, ProtoRequest protoRequest) {
        return null;
    }

    @Override
    public AuthorizationResponse getMustAuthenticateResponse(MustAuthenticateException cause) {
        return null;
    }

    @Override
    public ProtoResponse getMustAuthorizeResponse(MustAuthorizeException cause, ProtoRequest protoRequest) {
        return null;
    }

    @Override
    public AuthorizationResponse getMustAuthorizeResponse(MustAuthorizeException cause) {
        try {
            String userId = (String) cause.getDetails().getFirst().getValue();
            Optional<NamedValuedField> authContextIdField = cause.getDetails().stream().filter(f -> Objects.equals(f.getName(), "authContextId")).findFirst();

            OfficeExcelAttributes effectiveExcelAttributes;
            String state;

            if (authContextIdField.isPresent()) {
                String authContextId = (String) authContextIdField.get().getValue();
                effectiveExcelAttributes = officeExcelAttributeStore.load(authContextId);
                state = userId + "#" + authContextId;
            } else {
                effectiveExcelAttributes = officeExcelAttributes;
                state = userId;
            }

            OAuth20Service oAuth20Service = effectiveExcelAttributes.getOAuth20Service();
            String url = oAuth20Service.getAuthorizationUrl(state);
            String urlQueryParam = "&access_type=offline&approval_prompt=force&prompt=login";
            return new AuthorizationResponse(url + urlQueryParam, List.of());
        } catch (IOException ioException) {
            LOGGER.error("Failed to generate authorization URL: {}", ioException.getMessage(), ioException);
            throw new RuntimeException("Unable to generate Microsoft authorization URL. " +
                    "Please verify your extension configuration (Client ID, Client Secret, Redirect URI) and try again. " +
                    "Error details: " + ioException.getMessage(), ioException);
        }
    }
}
