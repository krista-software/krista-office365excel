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

package app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores;

import app.krista.extension.executor.Invoker;
import app.krista.extension.request.RoutingInfo;
import app.krista.extension.request.protos.http.HttpProtocol;
import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.extensions.util.KeyValueStore;
import org.jvnet.hk2.annotations.Service;

import javax.inject.Inject;
import java.io.IOException;
import java.util.Map;
import java.util.UUID;

@Service
public class OfficeExcelAttributeStore {

    private final Invoker invoker;
    private final KeyValueStore store;

    @Inject
    public OfficeExcelAttributeStore(Invoker invoker, KeyValueStore store) {
        this.invoker = invoker;
        this.store = store;
    }

    public OfficeExcelAttributes load(String authContextId) throws IOException {
        Map attributes = Constants.GSON.fromJson(String.valueOf(store.get(authContextId)), Map.class);
        return new OfficeExcelAttributes(
                (String) attributes.get(Constants.CLIENT_ID),
                (String) attributes.get(Constants.CLIENT_SECRET),
                (String) attributes.get(Constants.TENANT_ID),
                (String) attributes.get(Constants.EMAIL),
                invoker.getRoutingInfo().getRoutingURL(HttpProtocol.PROTOCOL_NAME, RoutingInfo.Type.APPLIANCE)
        );
    }

    public String save(OfficeExcelAttributes officeExcelAttributes) {
        String authContextId = UUID.randomUUID().toString();
        store.put(authContextId, Constants.GSON.toJson(
                Map.of(Constants.CLIENT_ID, officeExcelAttributes.getClientId(),
                        Constants.CLIENT_SECRET, officeExcelAttributes.getClientSecret(),
                        Constants.TENANT_ID, officeExcelAttributes.getTenantId(),
                        Constants.EMAIL, officeExcelAttributes.getEmailId())));

        return authContextId;
    }

    /**
     * Removes given authContext id from KeyValueStore
     *
     * @param authContextId authContext id
     */
    public void remove(String authContextId) {
        store.remove(authContextId);
    }
}
