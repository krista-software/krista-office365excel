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

import app.krista.extension.authorization.RequestAuthenticator;
import app.krista.extension.executor.Invoker;
import app.krista.extension.impl.anno.*;
import app.krista.extension.request.RoutingInfo;
import app.krista.extension.request.protos.http.HttpProtocol;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;

import javax.inject.Inject;
import java.io.IOException;
import java.util.Map;

@Java(version = Java.Version.JAVA_21)
@Field.Text(Constants.CLIENT_ID)
@Field.Text(value = Constants.CLIENT_SECRET, isSecured = true)
@Field.Text(value = Constants.TENANT_ID, isSecured = true)
@Field.Text(value = Constants.EMAIL)
@StaticResource(path = "docs", file = "docs")
@Extension(version = "2.0.28", jaxrsId = Constants.JAXRS_ID,  implementingDomainIds = "catEntryDomain_91825fb8-54b7-445e-90ea-a0f92b347a24", name = "Office 365 Excel")
public final class Office365ExcelExtension {
    private final OfficeExcelAttributes officeExcelAttributes;
    private final GraphServiceClientProviderFactory clientProviderFactory;
    private final OfficeExcelRequestAuthenticator requestAuthenticator;
    private final String routingId;

    @Inject
    public Office365ExcelExtension(Invoker invoker, OfficeExcelAttributes officeExcelAttributes,
                                   GraphServiceClientProviderFactory clientProviderFactory,
                                   OfficeExcelAttributeStore officeExcelAttributeStore) {
        this(invoker, officeExcelAttributes, clientProviderFactory,
                new OfficeExcelRequestAuthenticator(officeExcelAttributes, officeExcelAttributeStore));
    }

    public Office365ExcelExtension(Invoker invoker, OfficeExcelAttributes officeExcelAttributes,
                                   GraphServiceClientProviderFactory clientProviderFactory,
                                   OfficeExcelRequestAuthenticator requestAuthenticator) {
        this.officeExcelAttributes = officeExcelAttributes;
        this.clientProviderFactory = clientProviderFactory;
        this.requestAuthenticator = requestAuthenticator;
        this.routingId = invoker.getRoutingInfo().getRoutingURL(HttpProtocol.PROTOCOL_NAME, RoutingInfo.Type.APPLIANCE);
    }

    /**
     * This method is used for redirecting the MustAuthenticateException to RequestAuthenticator Interface.
     */
    @InvokerRequest(InvokerRequest.Type.AUTHENTICATOR)
    public RequestAuthenticator getRequestAuthenticator() {
        return this.requestAuthenticator;
    }

    @InvokerRequest(InvokerRequest.Type.CUSTOM_TABS)
    public Map<String, String> getCustomTabs() {
        return Map.of("Documentation", "static/docs");
    }

    @InvokerRequest(InvokerRequest.Type.VALIDATE_ATTRIBUTES)
    public void validate(Map<String, Object> attributes) throws IOException {
        OfficeExcelAttributes excelAttributes = OfficeExcelAttributes.create(routingId, attributes);
        testConnection(excelAttributes);
    }

    @InvokerRequest(InvokerRequest.Type.TEST_CONNECTION)
    public void testConnection() throws IOException {
        testConnection(officeExcelAttributes);
    }

    @InvokerRequest(InvokerRequest.Type.INVOKER_UPDATED)
    public void attributesUpdated(Map<String, Object> oldAttributes, Map<String, Object> newAttributes) {
        officeExcelAttributes.update(newAttributes);
    }

    private void testConnection(OfficeExcelAttributes excelAttributes) throws IOException {
        String email = excelAttributes.getEmailId();
        if (email != null && !email.isBlank()) {
            clientProviderFactory.create(excelAttributes)
                    .getGraphServiceClientForAdmin()
                    .me().buildRequest().get();
        }
    }
}
