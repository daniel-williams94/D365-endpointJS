const ENV_VAR_NAME = "Demo - Endpoint Addresses";
const PROGRESS_INDICATOR_MESSAGE = "Processing your request...";

/**
 * Main function to handle the demo process.
 * @param {Object} formContext - Context of the current form.
 * @param {string} id - GUID of the entity.
 */
async function demo(formContext, id) {
    if (!id) {
        logError("Invalid ID", new Error("ID is required for the demo process."));
        return;
    }

    const options = {
        method: "POST",
        headers: new Headers({ "Content-Type": "application/json" }),
        redirect: "follow",
        body: JSON.stringify({ id: cleanBraces(id) }),
    };

    Xrm.Utility.showProgressIndicator(PROGRESS_INDICATOR_MESSAGE);

    try {
        const endpoint = await getEndpointForCurrentEnvironment();
        if (!endpoint) {
            throw new Error("No endpoint found for the current environment.");
        }

        const response = await makeHttpRequest(endpoint, options);

        if (response.ok) {
            console.log("Request succeeded:", response);
        } else {
            throw new Error(`Request failed with status ${response.status}: ${response.statusText}`);
        }
    } catch (error) {
        logError("Error in demo function", error);
    } finally {
        Xrm.Utility.closeProgressIndicator();
    }
}

/**
 * Retrieves the endpoint URL for the current environment.
 * @returns {Promise<string>} - The endpoint URL.
 */
async function getEndpointForCurrentEnvironment() {
    const demoEndpointEnvVar = await getEnvironmentVariable(ENV_VAR_NAME);
    if (!demoEndpointEnvVar?.value) {
        throw new Error(`Invalid or missing environment variable: "${ENV_VAR_NAME}"`);
    }

    const currentEnvironment = Xrm.Utility.getGlobalContext().getClientUrl();
    const { Endpoints } = JSON.parse(demoEndpointEnvVar.value);

    return (
        Endpoints.find(ep => ep.environment === currentEnvironment)?.endpoint || null
    );
}

/**
 * Retrieves the value of an environment variable by its name.
 * @param {string} envVarName - The name of the environment variable.
 * @returns {Promise<Object>} - The environment variable value object.
 */
async function getEnvironmentVariable(envVarName) {
    try {
        const definitions = await getMultipleRecords(
            "environmentvariabledefinition",
            `?$select=environmentvariabledefinitionid,displayname&$filter=displayname eq '${envVarName}'`
        );

        if (definitions.entities.length === 1) {
            const definitionId = definitions.entities[0].environmentvariabledefinitionid;

            const values = await getMultipleRecords(
                "environmentvariablevalue",
                `?$select=value&$filter=_environmentvariabledefinitionid_value eq '${definitionId}'`
            );

            if (values.entities.length === 1) {
                return { value: values.entities[0].value };
            }
        }

        throw new Error(`Environment variable "${envVarName}" not found or invalid.`);
    } catch (error) {
        logError("Error retrieving environment variable", error);
        throw error;
    }
}

/**
 * Retrieves multiple records from a given table using a query.
 * @param {string} table - The table name.
 * @param {string} query - The query string.
 * @returns {Promise<Object>} - The retrieved records.
 */
async function getMultipleRecords(table, query) {
    try {
        return await Xrm.WebApi.online.retrieveMultipleRecords(table, query);
    } catch (error) {
        logError(`Error retrieving records from table "${table}"`, error);
        throw error;
    }
}

/**
 * Sends an HTTP request to the given endpoint with specified options.
 * @param {string} endpoint - The endpoint URL.
 * @param {Object} options - The request options.
 * @returns {Promise<Response>} - The HTTP response.
 */
async function makeHttpRequest(endpoint, options) {
    try {
        const response = await fetch(endpoint, options);

        if (!response.ok) {
            throw new Error(`HTTP error: ${response.status} - ${response.statusText}`);
        }

        return response;
    } catch (error) {
        logError("Error during HTTP request", error);
        throw error;
    }
}

/**
 * Removes curly braces from a GUID string and makes it lower case.
 * @param {string} guid - The GUID string.
 * @returns {string} - The cleaned GUID.
 */
function cleanBraces(guid) {
    if (!guid || typeof guid !== "string") {
        throw new Error("Invalid GUID provided.");
    }
    return guid.replace(/[\{\}]/g, "");
}

/**
 * Logs errors with a descriptive message.
 * @param {string} message - The error message.
 * @param {Error} error - The error object.
 */
function logError(message, error) {
    console.error(`${message}:`, error.message || error, error.stack || "");
}
