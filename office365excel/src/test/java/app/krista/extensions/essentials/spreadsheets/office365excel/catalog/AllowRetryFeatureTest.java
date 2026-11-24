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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog;

import app.krista.extension.executor.ExtensionResponse;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.handlers.AreaOperationHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.Validator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.ksdk.context.RequestContext;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.*;
import java.util.function.Function;
import java.util.function.Supplier;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * Unit tests for the Allow Retry feature in AreaOperationHandler.
 * 
 * Tests verify that:
 * 1. When allowRetry = TRUE and validation fails → generates interactive 3-step retry flow
 * 2. When allowRetry = FALSE/NULL and validation fails → returns direct error message
 * 3. When validation passes → executes operation regardless of allowRetry value
 */
@ExtendWith(MockitoExtension.class)
@DisplayName("Allow Retry Feature Tests")
class AllowRetryFeatureTest {

    @Mock
    private ValidationOrchestrator mockValidationOrchestrator;

    @Mock
    private ErrorHandlingStateManager mockStateManager;

    @Mock
    private ExtensionResponseGenerator mockResponseGenerator;

    @Mock
    private RequestContext mockRequestContext;

    @Mock
    private TelemetryHelper mockTelemetryHelper;

    @Mock
    private Supplier<String> mockOperation;

    @Mock
    private Function<String, Map<String, Object>> mockResultMapper;

    @Mock
    private Function<String, Map<String, Object>> mockStateMapper;

    private AreaOperationHandler<String> handler;

    @BeforeEach
    void setUp() {
        handler = new AreaOperationHandler<>(
                mockValidationOrchestrator,
                mockStateManager,
                mockResponseGenerator,
                mockRequestContext,
                mockTelemetryHelper
        );
    }

    @Test
    @DisplayName("When Allow Retry is TRUE and validation fails - Should generate confirmation response")
    void testAllowRetryEnabled_ValidationFails_GeneratesConfirmationResponse() {
        // Arrange
        String operationName = "deleteSheet";
        Map<Validator.ValidationResource, String> validationMap = new HashMap<>();
        validationMap.put(Validator.ValidationResource.FILE_ID, "invalid-file-id");

        Map<String, Object> stateValues = new HashMap<>();
        stateValues.put("fileId", "invalid-file-id");
        stateValues.put("sheetName", "Sheet1");

        String confirmationAction = SubCatalogConstants.CONFIRM_REENTER_DELETE_SHEET;
        Boolean allowRetry = Boolean.TRUE;

        // Create mock validation results
        List<ValidationOrchestrator.ValidationResult> validationResults = List.of(
                new ValidationOrchestrator.ValidationResult(
                        "File ID 'invalid-file-id' is invalid or does not exist",
                        "Please confirm if you want to re-enter the File ID",
                        "Enter the correct File ID",
                        "File Id",
                        "TEXT"
                )
        );

        ExtensionResponse expectedConfirmationResponse = mock(ExtensionResponse.class);

        // Mock validation orchestrator to return validation errors
        when(mockValidationOrchestrator.validate(validationMap)).thenReturn(validationResults);

        // Mock state mapper to return a state map
        Map<String, Object> expectedStateMap = Map.of("stateId", "test-state-id");
        when(mockStateMapper.apply(anyString())).thenReturn(expectedStateMap);

        // Mock response generator to return confirmation response
        when(mockResponseGenerator.generateConfirmationResponse(
                any(ExtensionResponse.Error.ExceptionType.class),
                anyList(),
                eq(confirmationAction),
                anyMap()
        )).thenReturn(expectedConfirmationResponse);

        // Act
        ExtensionResponse response = handler.execute(
                operationName,
                validationMap,
                stateValues,
                confirmationAction,
                mockStateMapper,
                mockOperation,
                mockResultMapper,
                Map.of("operation", "deleteSheet"),
                allowRetry
        );

        // Assert
        assertNotNull(response, "Response should not be null");
        assertEquals(expectedConfirmationResponse, response, "Should return confirmation response");

        // Verify validation was performed
        verify(mockValidationOrchestrator, times(1)).validate(validationMap);

        // Verify confirmation response was generated
        verify(mockResponseGenerator, times(1)).generateConfirmationResponse(
                any(ExtensionResponse.Error.ExceptionType.class),
                eq(validationResults),
                eq(confirmationAction),
                anyMap()
        );

        // Verify direct error response was NOT generated
        verify(mockResponseGenerator, never()).generateValidationErrorResponse(anyList());

        // Verify operation was NOT executed
        verify(mockOperation, never()).get();
    }

    @ParameterizedTest
    @CsvSource({
            "false, Allow Retry is FALSE",
            ", Allow Retry is NULL"
    })
    @DisplayName("When Allow Retry is FALSE/NULL and validation fails - Should return direct error response")
    void testAllowRetryDisabled_ValidationFails_ReturnsDirectError(Boolean allowRetry, String scenario) {
        // Arrange
        String operationName = "deleteSheet";
        Map<Validator.ValidationResource, String> validationMap = new HashMap<>();
        validationMap.put(Validator.ValidationResource.FILE_ID, "invalid-file-id");

        Map<String, Object> stateValues = new HashMap<>();
        stateValues.put("fileId", "invalid-file-id");
        stateValues.put("sheetName", "Sheet1");

        String confirmationAction = SubCatalogConstants.CONFIRM_REENTER_DELETE_SHEET;

        // Create mock validation results
        List<ValidationOrchestrator.ValidationResult> validationResults = List.of(
                new ValidationOrchestrator.ValidationResult(
                        "File ID 'invalid-file-id' is invalid or does not exist",
                        "Please confirm if you want to re-enter the File ID",
                        "Enter the correct File ID",
                        "File Id",
                        "TEXT"
                )
        );

        ExtensionResponse expectedDirectErrorResponse = mock(ExtensionResponse.class);

        // Mock validation orchestrator to return validation errors
        when(mockValidationOrchestrator.validate(validationMap)).thenReturn(validationResults);

        // Mock response generator to return direct error response
        when(mockResponseGenerator.generateValidationErrorResponse(validationResults))
                .thenReturn(expectedDirectErrorResponse);

        // Act
        ExtensionResponse response = handler.execute(
                operationName,
                validationMap,
                stateValues,
                confirmationAction,
                mockStateMapper,
                mockOperation,
                mockResultMapper,
                Map.of("operation", "deleteSheet"),
                allowRetry
        );

        // Assert
        assertNotNull(response, "Response should not be null for scenario: " + scenario);
        assertEquals(expectedDirectErrorResponse, response, "Should return direct error response for scenario: " + scenario);

        // Verify validation was performed
        verify(mockValidationOrchestrator, times(1)).validate(validationMap);

        // Verify direct error response was generated
        verify(mockResponseGenerator, times(1)).generateValidationErrorResponse(validationResults);

        // Verify confirmation response was NOT generated
        verify(mockResponseGenerator, never()).generateConfirmationResponse(
                any(), anyList(), anyString(), anyMap()
        );

        // Verify no state was saved
        verify(mockStateManager, never()).put(anyString(), anyMap());

        // Verify operation was NOT executed
        verify(mockOperation, never()).get();
    }

    @ParameterizedTest
    @CsvSource({
            "true, Allow Retry is TRUE",
            "false, Allow Retry is FALSE",
            ", Allow Retry is NULL"
    })
    @DisplayName("When validation passes - Should execute operation successfully regardless of allowRetry value")
    void testValidationPasses_ExecutesOperation(Boolean allowRetry, String scenario) {
        // Arrange
        Map<Validator.ValidationResource, String> validationMap = new HashMap<>();
        validationMap.put(Validator.ValidationResource.FILE_ID, "valid-file-id");

        String operationResult = "Operation completed successfully";
        Map<String, Object> resultMap = Map.of("result", operationResult);

        // Mock validation orchestrator to return no errors (empty list = validation passes)
        when(mockValidationOrchestrator.validate(validationMap)).thenReturn(Collections.emptyList());

        // Mock operation execution
        when(mockOperation.get()).thenReturn(operationResult);
        when(mockResultMapper.apply(operationResult)).thenReturn(resultMap);

        // Act
        ExtensionResponse response = handler.execute(
                "testOperation",
                validationMap,
                new HashMap<>(),
                "confirmAction",
                mockStateMapper,
                mockOperation,
                mockResultMapper,
                Map.of(),
                allowRetry
        );

        // Assert
        assertNotNull(response, "Response should not be null for scenario: " + scenario);

        // Verify validation was performed
        verify(mockValidationOrchestrator, times(1)).validate(validationMap);

        // Verify operation was executed
        verify(mockOperation, times(1)).get();

        // Verify result was mapped
        verify(mockResultMapper, times(1)).apply(operationResult);

        // Verify no error responses were generated
        verify(mockResponseGenerator, never()).generateValidationErrorResponse(anyList());
        verify(mockResponseGenerator, never()).generateConfirmationResponse(any(), anyList(), anyString(), anyMap());

        // Verify no state was saved
        verify(mockStateManager, never()).put(anyString(), anyMap());
    }
}

