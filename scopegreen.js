/**
 * Fetches LCA metrics data with caching to prevent timeouts.
 *
 * @param {string} itemName - Name of the item to find metrics for (max 100 chr)
 * @param {string} year - (Optional; Default: empty) Year associated with the item (≥2000)
 * @param {string} geography - (Optional; Default: empty) Geography associated with the item (max 50 chr)
 * @param {string} metric - (Optional; Default: Carbon footprint) Specific metric to search for (one of: Carbon footprint, EF3.1 Score, Land Use)
 * @param {string} domain - (Optional; Default: Materials & Products) Filter results by domain category (one of: Materials & Products, Processing, Transport, Energy, Direct emissions)
 * @param {number} numMatches - (Optional; Default: 1) Number of matches to return (1-3)
 * @param {string} mode - (Optional; Default: lite) Currently only lite available
 * @param {boolean} notEnglish - (Optional; Default: false) Set to true if the item_name is NOT in English and should be translated before search
 * @param {string} unit - (Optional; Default: empty) Target unit for the functional/base unit (denominator) for metric conversion (e.g., "g", "kWh", "lb"). Converts value and denominator unit, keeping numerator.
 * @return {Object[][]} The API response formatted as rows for display in a spreadsheet
 * @customfunction
 */
function SCOPEGREEN(itemName, year, geography, metric, domain, numMatches, mode, notEnglish, unit) {
    // For debugging parameter issues
    Logger.log("Parameters received:");
    Logger.log("itemName: " + itemName);
    Logger.log("year: " + year);
    Logger.log("geography: " + geography);
    Logger.log("metric: " + metric);
    Logger.log("domain: " + domain);
    Logger.log("numMatches: " + numMatches);
    Logger.log("mode: " + mode);
    Logger.log("notEnglish: " + notEnglish);
    Logger.log("unit: " + unit); // Added logging for unit
  
    // Basic preparation of parameters (no validation - the API will handle that)
    itemName = itemName ? itemName.toString().trim() : "";
    year = year ? year.toString().trim() : "";
    geography = geography ? geography.toString().trim() : "";
    metric = metric ? metric.toString().trim() : "Carbon footprint";
    domain = domain ? domain.toString().trim() : "";
    numMatches = numMatches ? parseInt(numMatches) : 1;
    mode = mode ? mode.toString().trim().toLowerCase() : "lite";
    notEnglish = notEnglish === true || notEnglish === "true" ? true : false;
    unit = unit ? unit.toString().trim() : ""; // Added preparation for unit
  
    // Create a cache key based on input parameters
    const cacheKey = `LCA_${itemName}_${year}_${geography}_${metric}_${domain}_${numMatches}_${mode}_${notEnglish}_${unit}`; // Added unit to cache key
  
    // Try to get cached result first
    const cache = CacheService.getUserCache();
    const cachedResult = cache.get(cacheKey);
  
    if (cachedResult) {
      try {
        return JSON.parse(cachedResult);
      } catch (e) {
        // If cache parse fails, continue with API call
        console.log("Cache parse failed:", e);
      }
    }
  
    // API configuration
    const API_BASE_URL = "https://scopegreen-main-1a948ab.d2.zuplo.dev/api/metrics/search";
    const API_KEY = "your-api-key";
  
    // Build the query parameters for the GET request
    let queryParams = `item_name=${encodeURIComponent(itemName)}`;
  
    // Add optional parameters if they have values
    if (year) queryParams += `&year=${encodeURIComponent(year)}`;
    if (geography) queryParams += `&geography=${encodeURIComponent(geography)}`;
    if (metric) queryParams += `&metric=${encodeURIComponent(metric)}`;
    if (mode) queryParams += `&mode=${encodeURIComponent(mode)}`;
    queryParams += `&web_mode=false`; // Always false as in the original
    if (numMatches) queryParams += `&num_matches=${encodeURIComponent(numMatches)}`;
    if (domain) queryParams += `&domain=${encodeURIComponent(domain)}`;
    queryParams += `&not_english=${encodeURIComponent(notEnglish)}`;
    if (unit) queryParams += `&unit=${encodeURIComponent(unit)}`; // Added unit to query parameters
  
    // Construct the full URL with query parameters
    const fullUrl = `${API_BASE_URL}?${queryParams}`;
    Logger.log("Request URL: " + fullUrl); // Log the final URL
  
    // API request options for GET
    const options = {
      "method": "get",
      "headers": {
        "Authorization": "Bearer " + API_KEY
      },
      "muteHttpExceptions": true,
      "timeout": 50000  // Set timeout to 50 seconds to stay within limits
    };
  
    try {
      // Make the API request with timeout control
      const response = UrlFetchApp.fetch(fullUrl, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
  
      // Log the raw response for debugging
      Logger.log("API Response Code: " + responseCode);
      Logger.log("API Response Text: " + responseText.substring(0, 500)); // Log first 500 chars
  
      // Check specifically for rate limit errors (HTTP 429)
      if (responseCode === 429) {
        // Rate limit exceeded
        try {
          const errorData = JSON.parse(responseText);
          if (errorData.error && errorData.error.message) {
            return [[`Rate limit exceeded: ${errorData.error.message}`]];
          }
        } catch (parseErr) {
          // If parsing fails, return a generic rate limit message
          return [["Rate limit exceeded: Please try again later."]];
        }
      }
  
      // Parse the JSON for successful responses
      let data;
      try {
        data = JSON.parse(responseText);
        // Log the parsed data structure
        Logger.log("Parsed data keys: " + Object.keys(data).join(", "));
      } catch (parseErr) {
        // Return the raw response text instead of an error message
        return [[responseText]];
      }
  
      // Handle error responses from the API
      if (data.error) {
        return [[`${data.error.code}: ${data.error.message}`]];
      }
  
      // FIXED: More robust check for "No match found" case
      if (data.message && typeof data.message === 'string' &&
          data.message.includes("No good match was found")) {
        return [[data.message]];
      }
  
      // Process successful response with matches
      if (data.matches && data.matches.length > 0) {
        const result = [];
  
        // Only process the available matches
        for (let i = 1; i <= Math.min(numMatches, data.matches.length); i++) {
          const match = data.matches.find(m => m.rank === i);
  
          if (match) {
            // Match name with rank
            result.push("Match " + i + ": " + match.matched_name);
  
            // Metric value (numeric only)
            result.push(match.metric.value);
  
            // Unit (text only)
            result.push(match.metric.unit);
  
            // Year
            result.push(match.year || "");
  
            // Geography
            result.push(match.geography || "");
  
            // Source - Get name and URL from separate fields
            let sourceName = match.source || "";      // Get name from 'source' field
            let sourceUrl = match.source_link || ""; // Get URL from 'source_link' field
  
            // Add source name and URL separately
            result.push(sourceName); // Column 6
            result.push(sourceUrl);  // Column 7
  
            // Conversion info (if present)
            result.push(match.conversion_info || ""); // Added conversion_info
          }
        }
  
        // Add full explanation without shortening
        if (data.explanation) {
          result.push(data.explanation);
        } else {
          result.push("");
        }
  
        // Cache the result for future calls (max 120 seconds / 2 minutes)
        try {
          cache.put(cacheKey, JSON.stringify([result]), 120);
        } catch (cacheErr) {
          console.log("Cache storage error:", cacheErr);
        }
  
        return [result];
      }
  
      // If we get here, log what we actually received for debugging
      Logger.log("Unexpected response structure: " + JSON.stringify(data));
  
      // Check again for ANY message field as a final fallback
      if (data.message) {
        return [[`Message from API: ${data.message}`]];
      }
  
      // Return the raw response as a last resort
      return [[responseText]];
    } catch (error) {
      // Only keep timeout handling since it's client-side
      if (error.toString().includes("timeout") || error.toString().includes("execution time")) {
        return [["Error: API request timed out. Try simplifying your query."]];
      }
  
      // For all other errors, return the error message as is
      return [["Error: " + error.toString()]];
    }
  }