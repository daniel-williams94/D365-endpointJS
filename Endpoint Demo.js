async function demo(formContext, id) {
    const headers = new Headers({ "Content-Type": "application/json" });
    const options = {
        method: "POST",
        headers,
        redirect: "follow",
        body: JSON.stringify({ id: cleanBraces(id) }),
    };

    const progressIndicatorMessage = "Processing your request...";
    Xrm.Utility.showProgressIndicator(progressIndicatorMessage);

    try {
        const endpoint = await getEndpointForCurrentEnvironment();
        if (!endpoint) {
            throw new Error("Endpoint not found for the current environment.");
        }

        const response = await makeHttpRequest(endpoint, options);

        if (response.ok) {
            const responseData = response;
            console.log("Request succeeded:", responseData);
        } else {
            throw new Error(`Request failed with status ${response.status}: ${response.statusText}`);
        }
    } catch (error) {
        logError("Demo function error", error);
    } finally {
        Xrm.Utility.closeProgressIndicator();
    }
}

async function getEndpointForCurrentEnvironment() {
    const envVarName = "Demo - Endpoint Addresses";

    const demoEndpointEnvVar = await getEnvironmentVariable(envVarName);
    if (!demoEndpointEnvVar?.value) {
        throw new Error(`Invalid or missing environment variable: "${envVarName}"`);
    }

    const currentEnvironment = Xrm.Utility.getGlobalContext().getClientUrl();
    const { Endpoints } = JSON.parse(demoEndpointEnvVar.value);

    return Endpoints.find(ep => ep.environment === currentEnvironment)?.endpoint || null;
}

async function getEnvironmentVariable(envVarName) {
    try {
        const result = await getMultipleRecords(
            "environmentvariabledefinition",
            `?$select=environmentvariabledefinitionid,displayname&$filter=displayname eq '${envVarName}'`
        );

        if (result.entities.length === 1) {
            const definitionId = result.entities[0].environmentvariabledefinitionid;

            const valueResult = await getMultipleRecords(
                "environmentvariablevalue",
                `?$select=value&$filter=_environmentvariabledefinitionid_value eq '${definitionId}'`
            );

            if (valueResult.entities.length === 1) {
                return { value: valueResult.entities[0].value };
            }
        }

        throw new Error(`Environment variable "${envVarName}" not found.`);
    } catch (error) {
        logError("Error retrieving environment variable", error);
        throw error;
    }
}

async function getMultipleRecords(table, query) {
    try {
        return await Xrm.WebApi.online.retrieveMultipleRecords(table, query);
    } catch (error) {
        logError(`Error retrieving records from table: "${table}"`, error);
        throw error;
    }
}

async function makeHttpRequest(endpoint, options) {
    try {
        const response = await fetch(endpoint, options);

        if (!response.ok) {
            throw new Error(`HTTP error: ${response.status} - ${response.statusText}`);
        }

        return response;
    } catch (error) {
        logError("HTTP request error", error);
        throw error;
    }
}

function cleanBraces(guid) {
    if (!guid || typeof guid !== "string") {
        throw new Error("Invalid GUID provided to cleanBraces.");
    }
    return guid.replace(/[\{\}]/g, '');
}

function logError(message, error) {
    console.error(`${message}:`, error.message || error, error.stack || "");
}