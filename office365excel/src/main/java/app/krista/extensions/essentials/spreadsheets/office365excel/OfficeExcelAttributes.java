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

import app.krista.extension.executor.Invoker;
import app.krista.extension.request.RoutingInfo;
import app.krista.extension.request.protos.http.HttpProtocol;
import app.krista.extension.util.InvokerAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import com.github.scribejava.apis.MicrosoftAzureActiveDirectory20Api;
import com.github.scribejava.core.builder.ServiceBuilder;
import com.github.scribejava.core.oauth.OAuth20Service;
import org.jvnet.hk2.annotations.Service;

import javax.inject.Inject;
import java.util.Map;

@Service
public class OfficeExcelAttributes {

    private String clientId;
    private String clientSecret;
    private String tenantId;
    private final String callbackUrl;
    public String emailId;
    private OAuth20Service oAuth20Service;

    public static OfficeExcelAttributes create(String routingId, Map<String, Object> attributes) {
        return new OfficeExcelAttributes(attributes, routingId);

    }

    @Inject
    public OfficeExcelAttributes(Invoker invoker) {
        this(invoker.getAttributes(),
                invoker.getRoutingInfo().
                        getRoutingURL(HttpProtocol.PROTOCOL_NAME, RoutingInfo.Type.APPLIANCE));

    }

    public OfficeExcelAttributes(Map<String, Object> attributes, String routingId) {
        this(InvokerAttributes.getStringOrNull(attributes, Constants.CLIENT_ID),
                InvokerAttributes.getStringOrNull(attributes, Constants.CLIENT_SECRET),
                InvokerAttributes.getStringOrNull(attributes, Constants.TENANT_ID),
                InvokerAttributes.getStringOrNull(attributes, Constants.EMAIL), routingId);
    }

    public OfficeExcelAttributes(String clientId, String clientSecret, String tenantId, String emailId, String routingId) {
        this.clientId = clientId;
        this.clientSecret = clientSecret;
        this.tenantId = tenantId;
        this.emailId = emailId;
        this.callbackUrl = routingId + "/" + Constants.JAXRS_ID + "/" + Constants.CALLBACK;
    }

    public void update(Map<String, Object> newAttributes) {
        clientId = InvokerAttributes.getStringOrNull(newAttributes, Constants.CLIENT_ID);
        clientSecret = InvokerAttributes.getStringOrNull(newAttributes, Constants.CLIENT_SECRET);
        tenantId = InvokerAttributes.getStringOrNull(newAttributes, Constants.TENANT_ID);
        emailId = InvokerAttributes.getStringOrNull(newAttributes, Constants.EMAIL);
    }

    public String getClientId() {
        return clientId;
    }

    public String getClientSecret() {
        return clientSecret;
    }

    public String getTenantId() {
        return tenantId;
    }

    public String getEmailId() {
        return emailId;
    }

    public String getCallbackUrl() {
        return callbackUrl;
    }

    public OAuth20Service getOAuth20Service() {
        if (oAuth20Service == null) {
            oAuth20Service = new ServiceBuilder(clientId)
                    .apiSecret(clientSecret)
                    .defaultScope(Constants.SCOPE)
                    .callback(getCallbackUrl())
                    .build(MicrosoftAzureActiveDirectory20Api.custom(tenantId));
        }
        return oAuth20Service;
    }
}
