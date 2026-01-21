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

package app.krista.extensions.essentials.spreadsheets.office365excel.util.constants;

import app.krista.model.base.FreeForm;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;

import java.util.Map;

public final class Constants {

    private Constants() {
    }

    public static final Gson GSON = new GsonBuilder()
            .serializeSpecialFloatingPointValues()
            .setPrettyPrinting()
            .setLenient()  // Allow lenient parsing to handle malformed JSON from Microsoft Graph API
            .create();

    public static final String STATE = "state";

    public static final String SPACE = " ";

    public static final String BACKWARD_SLASH = "/";

    public static final String FILES_READWRITE_ALL = "Files.ReadWrite.All";

    public static final String FILES_READWRITE = "Files.ReadWrite";

    public static final String SITES_READ_ALL = "Sites.Read.All";

    public static final String EMPTY_STRING = "";


    public static final Character A = 'A';

    public static final String SHEET_RANGE_SEPARATOR = "!";

    public static final String COLON = ":";

    public static final String COMMA = ",";

    public static final String CELL_ID = "Cell ID";

    public static final String FORMULA = "Formula";

    public static final String VALUE = "Value";

    public static final String USER_READ_WRITE = "User.ReadWrite";


    public static final String FIRST_COLUMN = "A";

    public static final String LAST_COLUMN = "XFD";

    public static final String EMAIL = "email";

    public static final String CLIENT_ID = "clientID";

    public static final String CLIENT_SECRET = "clientSecret";

    public static final String ID = "id";

    public static final String EXCEL_FILE_EXTENSION = ".xlsx";

    public static final String OPENID = "openid";

    public static final String OFFLINE_ACCESS = "offline_access";

    public static final String CODE = "code";

    public static final String JAXRS_ID = "excel365extn";

    public static final String SCOPE = OPENID + SPACE
            + OFFLINE_ACCESS + SPACE
            + FILES_READWRITE
            + SPACE + SITES_READ_ALL
            + SPACE + FILES_READWRITE_ALL
            + SPACE + USER_READ_WRITE;

    public static final String CALLBACK = "callback";

    public static final String TENANT_ID = "tenantId";

    public static final String AUTHORITY = "https://login.microsoftonline.com/";

    public static final String MEDIA_URL = "http://platform:9991";

    public static final String MAX_VALUE = "Max_Value";

    public static final String MAX_VALUE_ROW_INDEX = "Max_Value_Row_Index";

    public static final String MIN_VALUE = "Min_Value";

    public static final String MIN_VALUE_ROW_INDEX = "Min_Value_Row_Index";

    public static final String DOWN = "Down";

    public static final String INVALID_RANGE_IS_GIVEN = "Invalid Excel range format: ";

    public static final String INVALID_COLUMN_RANGE_IS_GIVEN = "Invalid column range format: ";

    public static final String INVALID_CELL_ID_IS_GIVEN = "Invalid cell reference: ";

    public static final String WORKBOOK_WORKSHEETS = "/workbook/worksheets/";

    public static final String RANGE_ADDRESS = "/range(address='";

    public static final String SHIFT = "shift";

    public static final String ROW_DATA_NOT_PROVIDED = "No row data provided. Please provide either 'Rows' or 'Rows from CSV' to add data to the sheet.";

    public static final String DELETE = "')/delete";

    public static final String COL = ":";

    public static final String DATA = "Data";

    public static final String LEFT = "Left";

    public static final String UP = "Up";

    public static final String REFRESH_TOKEN_EXPIRED_PLEASE_REAUTHORIZE_YOURSELF = "Your authentication session has expired. Please re-authorize the connection to continue.";

    public static final String ERROR = "Error";

    public static final String TEXT = "Text";

    public static final String GRAPH_MICROSOFT_COM_V_1_0_ME = "https://graph.microsoft.com/v1.0/me";

    public static final FreeForm FREE_FORM = new FreeForm();

    public enum ApplyTo {All, Formats, Contents}

    public static final String AUTHORIZATION_PROMPT = "Authorization required. Please click 'Validate Attributes' to authenticate before saving your configuration.";

    public static final String USER_ID = "userId";

    public static final String AUTH_CONTEXT_ID = "authContextId";

    public static final String UTC = "UTC";

    public static final String MAILBOX_SETTINGS = "mailboxSettings";

    /**
     * Mapping of Windows timezone names to Java IANA timezone IDs.
     * Used to convert Microsoft Graph API timezone to Java TimeZone.
     */
    public static final Map<String, String> WINDOWS_TO_IANA = Map.ofEntries(
            Map.entry("India Standard Time", "Asia/Kolkata"),
            Map.entry("Pacific Standard Time", "America/Los_Angeles"),
            Map.entry("Mountain Standard Time", "America/Denver"),
            Map.entry("Central Standard Time", "America/Chicago"),
            Map.entry("Eastern Standard Time", "America/New_York"),
            Map.entry("US Mountain Standard Time", "America/Phoenix"),
            Map.entry("Alaskan Standard Time", "America/Anchorage"),
            Map.entry("Hawaiian Standard Time", "Pacific/Honolulu"),
            Map.entry("GMT Standard Time", "Europe/London"),
            Map.entry("W. Europe Standard Time", "Europe/Berlin"),
            Map.entry("Central Europe Standard Time", "Europe/Budapest"),
            Map.entry("Romance Standard Time", "Europe/Paris"),
            Map.entry("E. Europe Standard Time", "Europe/Bucharest"),
            Map.entry("FLE Standard Time", "Europe/Kiev"),
            Map.entry("GTB Standard Time", "Europe/Athens"),
            Map.entry("Russian Standard Time", "Europe/Moscow"),
            Map.entry("China Standard Time", "Asia/Shanghai"),
            Map.entry("Tokyo Standard Time", "Asia/Tokyo"),
            Map.entry("Korea Standard Time", "Asia/Seoul"),
            Map.entry("Singapore Standard Time", "Asia/Singapore"),
            Map.entry("W. Australia Standard Time", "Australia/Perth"),
            Map.entry("AUS Eastern Standard Time", "Australia/Sydney"),
            Map.entry("E. Australia Standard Time", "Australia/Brisbane"),
            Map.entry("West Asia Standard Time", "Asia/Tashkent"),
            Map.entry("Central Asia Standard Time", "Asia/Almaty"),
            Map.entry("SE Asia Standard Time", "Asia/Bangkok"),
            Map.entry("Myanmar Standard Time", "Asia/Rangoon"),
            Map.entry("Sri Lanka Standard Time", "Asia/Colombo"),
            Map.entry("Nepal Standard Time", "Asia/Kathmandu"),
            Map.entry("Bangladesh Standard Time", "Asia/Dhaka"),
            Map.entry("Pakistan Standard Time", "Asia/Karachi"),
            Map.entry("Arabian Standard Time", "Asia/Dubai"),
            Map.entry("Iran Standard Time", "Asia/Tehran"),
            Map.entry("Afghanistan Standard Time", "Asia/Kabul"),
            Map.entry("Israel Standard Time", "Asia/Jerusalem"),
            Map.entry("Arabic Standard Time", "Asia/Baghdad"),
            Map.entry("Turkey Standard Time", "Europe/Istanbul"),
            Map.entry("SA Pacific Standard Time", "America/Bogota"),
            Map.entry("SA Western Standard Time", "America/La_Paz"),
            Map.entry("SA Eastern Standard Time", "America/Cayenne"),
            Map.entry("Argentina Standard Time", "America/Buenos_Aires"),
            Map.entry("E. South America Standard Time", "America/Sao_Paulo"),
            Map.entry("Venezuela Standard Time", "America/Caracas"),
            Map.entry("Atlantic Standard Time", "America/Halifax"),
            Map.entry("Central America Standard Time", "America/Guatemala"),
            Map.entry("Canada Central Standard Time", "America/Regina"),
            Map.entry("Mexico Standard Time", "America/Mexico_City"),
            Map.entry("South Africa Standard Time", "Africa/Johannesburg"),
            Map.entry("Egypt Standard Time", "Africa/Cairo"),
            Map.entry("W. Central Africa Standard Time", "Africa/Lagos"),
            Map.entry("E. Africa Standard Time", "Africa/Nairobi"),
            Map.entry("New Zealand Standard Time", "Pacific/Auckland"),
            Map.entry("Fiji Standard Time", "Pacific/Fiji"),
            Map.entry("UTC", "UTC")
    );

}
