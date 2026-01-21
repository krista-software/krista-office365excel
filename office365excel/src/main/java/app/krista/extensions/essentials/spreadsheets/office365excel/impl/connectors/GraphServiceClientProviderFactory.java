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

import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.RefreshTokenStore;
import app.krista.ksdk.context.AuthorizationContext;
import app.krista.ksdk.context.RequestContext;
import org.jvnet.hk2.annotations.Service;

import javax.inject.Inject;
import java.io.IOException;

@Service
public class GraphServiceClientProviderFactory {

    private final RefreshTokenStore refreshTokenStore;
    private final OfficeExcelAttributeStore officeExcelAttributeStore;
    private final RequestContext requestContext;
    private final AuthorizationContext authorizationContext;
    private final GraphServiceClientProvider defaultClientProvider;


    @Inject
    public GraphServiceClientProviderFactory(RefreshTokenStore refreshTokenStore, OfficeExcelAttributeStore officeExcelAttributeStore,
                                             OfficeExcelAttributes officeExcelAttributes, RequestContext requestContext,
                                             AuthorizationContext authorizationContext) {
        this(refreshTokenStore, officeExcelAttributeStore, requestContext, authorizationContext,
                new GraphServiceClientProvider(refreshTokenStore, officeExcelAttributes, requestContext, authorizationContext, null));

    }

    public GraphServiceClientProviderFactory(RefreshTokenStore refreshTokenStore, OfficeExcelAttributeStore officeExcelAttributeStore,
                                             RequestContext requestContext, AuthorizationContext authorizationContext,
                                             GraphServiceClientProvider defaultClientProvider) {
        this.refreshTokenStore = refreshTokenStore;
        this.officeExcelAttributeStore = officeExcelAttributeStore;
        this.requestContext = requestContext;
        this.authorizationContext = authorizationContext;
        this.defaultClientProvider = defaultClientProvider;
    }

    public GraphServiceClientProvider create(OfficeExcelAttributes officeExcelAttributes) throws IOException {
        String authContextId = this.officeExcelAttributeStore.save(officeExcelAttributes);
        return new GraphServiceClientProvider(refreshTokenStore, officeExcelAttributes, requestContext, authorizationContext, authContextId);

    }

    public GraphServiceClientProvider create() {
        return defaultClientProvider;
    }

}