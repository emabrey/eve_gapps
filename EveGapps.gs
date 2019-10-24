/*

EVE GApps Script v1.0.0

Copyright 2019 Emily Mabrey

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

/*
Thanks to Kevin McDonald (a/k/a sudorandom) for the functions from the
evepraisal-google-sheets project.

Please check out his project: https://github.com/evepraisal/evepraisal-google-sheets.

My in-game name is Emilia Felinati - please let me know that you used my script. 
I love to learn I made something useful!
*/

/*jshint 
  esversion: 6,
  curly: true,
  strict: true,
  bitwise: true,
  eqeqeq: false,
  futurehostile: true,
  noarg: true,
  nocomma: true,
  nonbsp: true,
  undef: true,
  varstmt: false
*/

/*
  globals AuthMode, AuthorizationStatus, Browser, Button, ButtonSet, console,
  CacheService, CardService, Charset, ChartHiddenDimensionStrategy, ChartMergeStrategy,
  ChartType, Charts, ColorType, ColumnType, ComposedEmailType, ContentService, ContentType,
  CurveStyle, DigestAlgorithm, EventType, HtmlService, Icon, ImageStyle, InstallationSource,
  Jdbc, LinearOptimizationService, LoadIndicator, LockService, Logger, MacAlgorithm, MailApp,
  MatchType, MimeType, Month, OnClose, OpenAs, Orientation, PickerValuesLayout, PointStyle,
  Position, PropertiesService, RsaAlgorithm, SandboxMode, ScriptApp, SelectionInputType,
  Session, Status, TextButtonStyle, TriggerSource, UpdateDraftBodyType, UrlFetchApp,
  VariableType, Weekday, XFrameOptionsMode, XmlService
*/

/**
 * The various possible values for EVEPRAISAL formula
 */
var EVEPRAISAL = {

  /**
   * The available markets which Evepraisal can use to calculate an order's value at the specified market.
   * @see {@link DEFAULT#MARKET}
   */
  MARKETS: {
    universe: "universe",
    jita: "jita",
    amarr: "amarr",
    dodixie: "dodixie",
    hek: "hek",
    rens: "rens"
  },

  /**
   * The available order types.
   * @see {@link DEFAULT#ORDER_TYPE}
   */
  ORDER_TYPE: {
    sell: "sell",
    buy: "buy"
  },

  /**
   * The available order attributes. These are used to control the information retrieved/displayed about a particular order.
   * @see {@link DEFAULT#ATTRIBUTE}
   */
  ATTRIBUTE: {
    average: "avg",
    maximum: "max",
    median: "median",
    minimum: "min",
    percentile: "percentile",
    standard_deviation: "stddev",
    volume: "volume",
    order_count: "order_count"
  }
};

/**
 * The various default values used for option parameters in the various functions
 */
var DEFAULT = {

  /**
   * @type {string} MARKET The default market used by functions requiring an optionally specified market
   * @see {@link EVEPRAISAL#MARKET}
   */
  MARKET: "jita",

  /**
   * @type {string} ORDER_TYPE The default order type (buy or sell) used by functions requiring an optionally specified order type
   * @see {@link EVEPRAISAL#ORDER_TYPE}
   */
  ORDER_TYPE: "sell",

  /**
   * @type {string} ATTRIBUTE_SELL The default order attribute (min or max) used by functions requiring an optionally specified attribute for a sell order
   * @see {@link EVEPRAISAL#ATTRIBUTE}
   */
  ATTRIBUTE_SELL: "min",

  /**
   * @type {string} ATTRIBUTE_BUY The default order attribute (min or max) used by functions requiring an optionally specified attribute for a buy order
   * @see {@link EVEPRAISAL#ATTRIBUTE}
   */
  ATTRIBUTE_BUY: "max",

  /**
   * @type {string} LINKED_DESCRIPTION_TYPE The default site (eveInfo or zKillboard) for generated URLs that links to an item's description on a 3rd party site
   */
  LINKED_DESCRIPTION_TYPE: "zKillboard",

  /**
   * @type {string} FETCHED_DESCRIPTION_TYPE The default site (eveInfo or zKillboard) for fetching a 3rd party provided item description
   */
  FETCHED_DESCRIPTION_TYPE: "thonky",

  /**
   * @type {string} CONTENTTYPE The default content type (byte or text) for caching fetched URLs
   */
  CONTENTTYPE: "byte",

  /**
   * @type {string} EXPIRYTYPE The default expiry type (short or long) for caching fetched URLs
   */
  EXPIRYTYPE: "short",

  /**
   * @type {number} TIMEOUT The default timeout for for storing short-expiry cached values
   */
  TIMEOUT: 300,

  /**
   * @type {number} LONGTIMEOUT The default timeout for for storing long-expiry cached values
   */
  LONGTIMEOUT: 86400,

  /**
   * @type {array} URLPARAMS The default URL fetch parameters, as utilized via {@link UrlFetchApp#fetch(url, params)}
   */
  URLPARAMS: {},

  /**
   * @type {number} IMAGE_DIMENSION The default dimension (32 or 64) of the square images used for icons
   *
   */
  IMAGE_DIMENSION: 32
};

/**
 * @type {object} SITES The various site URL values used
 */
var SITES = {
  /**
   * @type {string} ZKILLBOARD_ITEMS_PREFIX The ZKillboard site's URL prefix for item related information
   */
  ZKILLBOARD_ITEMS_PREFIX: "https://zkillboard.com/item/",

  /**
   * @type {string} EVEINFO_ITEMS_PREFIX The EveInfo site's URL prefix for item related information
   */
  EVEINFO_ITEMS_PREFIX: "https://eveinfo.com/item/",

  /**
   * @type {string} THONKY_ITEMS_PREFIX The Thonky site's URL prefix for item related information
   */
  THONKY_ITEMS_PREFIX: "https://www.thonky.com/eve-online-guide/eve-item-database?typeID=",

  /**
   * @type {string} EVEPRAISAL_APPRAISAL_PREFIX The Evepraisal site's URL prefix for order appraisal
   */
  EVEPRAISAL_APPRAISAL_PREFIX: "https://evepraisal.com/a/",

  /**
   * @type {string} EVEPRAISAL_ITEM_PREFIX The Evepraisal site's URL prefix for item related information
   */
  EVEPRAISAL_ITEM_PREFIX: "https://evepraisal.com/item/",

  /**
   * @type {string} FUZZWORKS_ITEMID_PREFIXThe Fuzzworks site's URL prefix for the itemName to itemID API
   */
  FUZZWORKS_ITEMID_PREFIX: "https://www.fuzzwork.co.uk/api/typeid.php?typename=",

  /**
   * @type {string} EVEONLINE_IMAGESERVER_PREFIX The official EVE Online URL prefix for fetching official icons
   */
  EVEONLINE_IMAGESERVER_PREFIX: "https://imageserver.eveonline.com/Type/"

};

/**
 * Uses the {@link UrlFetchApp} to fetch the given URL using the provided parameters. The response content
 * is transformed into a series of bytes by default, or as a text value if contentType is "text".
 * The transformed response content is cached twice in a short and long lived cache.
 *
 * @param {string} url The URL of the content to fetch
 * @param {string} [contentType="byte"] Whether to return the content as text ("text") or as a byte[].
 * @param {number} [timeout=300] The short-expiry cache timeout
 * @param {number} [longTimeout=86400] The long-expiry cache timeout
 * @param {array} [urlParams={}] An array containing the URL parameters as specified in the {@link UrlFetchApp#fetch(url, params)} method
 * @return {byte[]\|string} The content of the fetched site
 */
function fetchCachedUrlContent(url, contentType, timeout, longTimeout, urlParams) {
  "use strict";

  requireNonNullParameter(url, "url");

  contentType = contentType == null ? DEFAULT.CONTENTTYPE : contentType;
  timeout = timeout == null ? DEFAULT.TIMEOUT : timeout;
  longTimeout = longTimeout == null ? DEFAULT.LONGTIMEOUT : longTimeout;
  urlParams = urlParams == null ? DEFAULT.URLPARAMS : urlParams;

  const CACHE_KEY = getCacheKey(url, contentType, "short");
  const CACHE_LONGTIMEOUT_KEY = getCacheKey(url, contentType, "long");

  const cache = CacheService.getScriptCache();

  const cachedValue = cache.get(CACHE_KEY);
  const cachedValueExists = cachedValue != null;

  if(cachedValueExists) {
    return cachedValue;
  }

  var data;

  try {
    data = UrlFetchApp.fetch(url, urlParams);
    data = contentType == "text" ? data.getContentText() : data.getContent();
  } catch (ex) {
    if(/Service invoked too many times for one day/gi.test(ex.toString())) {
      const longExpiryCacheValue = cache.get(CACHE_LONGTIMEOUT_KEY);
      const longExpiryCacheValueExists = longExpiryCacheValue != null;
      if(longExpiryCacheValueExists) {
        return longExpiryCacheValue;
      }
    }
  }

  try {
    cache.put(CACHE_KEY, data, timeout);
    cache.put(CACHE_LONGTIMEOUT_KEY, data, longTimeout);
  } catch (ex) {
    console.log("There was a caching problem: %s", ex.toString());
  }

  return data;
}

/**
 * Generates repeated calls to the specified function based upon the two-dimensional array present in the first element
 * in the given array of caller arguments. This function is a drop-in solution for implementing support for ranges
 * in a custom function that doesn't normally support them.
 *
 * @example 
 * //How to add support for range to function simpleFunc(arg1, arg2, arg3)
 * if(arguments[0] && arguments[0].map) {
 *   return multipleValueRangesHandler(simpleFunc, 3, arguments);
 * }
 * @param {object} function_obj The specified function which will be called
 * @param {number} ideal_arg_count The number of arguments the specified function has
 * @param {array} callerArguments The un-mangled arguments given to the custom function by the user
 * @return {array} The results of n-many calls to the specified function
 */
function multipleValueRangesHandler(function_obj, ideal_arg_count, callerArguments) {

  "use strict";

  for(var a = 0; a < callerArguments.length; a++) {
    if(callerArguments[a] && callerArguments[a].map && a != 0) {
      throw "Do not call " + function_obj.name + " with multiple ranges, use one range composed of consecutive columns";
    }
  }

  var arrayString = callerArguments[0].toString();

  var initialRegexpString = "";

  for(var b = 0; b < (ideal_arg_count - 1); b++) {
    initialRegexpString = initialRegexpString.concat("[^,]+,");
  }

  initialRegexpString = initialRegexpString.concat("[^,]+");

  var arrayStringSplit = arrayString.match(new RegExp(initialRegexpString, "g"));
  //splits into strings with the right number of args

  var numResults = arrayStringSplit.length;

  var results = Array(numResults);

  for(var c = 0; c < results.length; c++) {
    var currentArgsUnsplit = arrayStringSplit[c];

    var currentArgsSplit = currentArgsUnsplit.match(/[^,\n]+/g); //splits things into 1 string per arg

    for(var d = 0; d < currentArgsSplit.length; d++) {
      if(currentArgsSplit[d] == "") {
        currentArgsSplit[d] = null;
      }
    }

    results[c] = function_obj.apply(undefined, currentArgsSplit);
  }

  return results;
}

/**
 * Throws an exception with optional argument identifier if the provided argument is null, undefined or an empty string.
 * The check is performed in this order: (null || undefined || empty string) so as to prevent confusing results with strings.
 *
 * @param {*} arg The object which we are requiring to meet the constraints on nullity and emptiness
 * @param {string} [argName=""] A name for the verified object, such as the variable name, used to generate meaningful error messages
 */
function requireNonNullParameter(arg, argName) {
  "use strict";

  argName = argName == null ? "" : argName;

  if(arg == null || arg == undefined || arg == "") {
    throw "required parameter " + argName + " not given";
  }
}

/**
 * Generates the CacheService key used to cache the given URL using the given content type and expiration length. Note that
 * this method does not care about the actual length of a "short" cache term or a "long" cache term.
 *
 * @param {string} url The url of the content to cache
 * @param {string} [contentType="byte"] Whether to cache the content as text ("text") or as a byte[] ("byte")
 * @param {string} [expiryType="short"] The type of cache expiration: long-term expiry or short-term expiry
 * @return {string} A cache key for the given value and metadata which does not exceed the maximum of 250 characters
 */
function getCacheKey(url, contentType, expiryType) {
  "use strict";

  requireNonNullParameter(url, "url");

  contentType = contentType == null ? DEFAULT.CONTENTTYPE : contentType;
  expiryType = expiryType == null ? DEFAULT.EXPIRYTYPE : expiryType;

  const CACHE_CONTENT_PREFIX = contentType == "text" ? "text:" : "byte:";
  const CACHE_TIMEOUT_PREFIX = expiryType == "long" ? "long:" : "";

  const CACHE_KEY_OVERSIZED = CACHE_TIMEOUT_PREFIX + CACHE_CONTENT_PREFIX + url;

  const MAX_KEY_LENGTH = 250;

  const CACHE_KEY = CACHE_KEY_OVERSIZED.substring(0, MAX_KEY_LENGTH);

  return CACHE_KEY;
}

/**
 * Extracts the description content from the source of an entire item webpage for the EveInfo service
 * 
 * @param {string} htmlResponse The HTML content of the previously fetched web page
 * @returns {string} The extracted description of the item extracted from within the provided HTML content
 */
function extractDescriptionFromEveInfoResponse(htmlResponse) {
  "use strict";

  requireNonNullParameter(htmlResponse, "htmlResponse");

  var entireTag = htmlResponse.match(/<\s*p class="desc" itemprop="articleBody"[^>]*>(.*?)<\s*\/\s*p>/g)[0];

  var startPos = 0 + '<p class="desc" itemprop="articleBody">'.length;
  var endPos = entireTag.length - "</p>".length;

  startPos = startPos < 0 ? 0 : startPos;
  endPos = endPos < 0 ? 0 : endPos;

  var tagContents = entireTag.substring(startPos, endPos);

  return tagContents;
}

/**
 * Extracts the description content from the source of an entire item webpage for the zKillboard
 * service.
 * 
 * Currently this is unimplemented, due to the zKillboard site not being easy to parse.
 * As a result calling this method will simply throw.
 * 
 * @param {string} htmlResponse The HTML content of the previously fetched web page
 * @returns {string} The extracted description of the item extracted from within the provided HTML content
 */
function extractDescriptionFromZKillboardResponse(htmlResponse) {
  "use strict";

  htmlResponse = htmlResponse; //This function deliberately left unimplemented
  throw "Currently ZKillboard is not supported for extracting description strings";
}

/**
 * Extracts the description content from the source of an entire item webpage for the Thonky service
 * 
 * @param {string} htmlResponse The HTML content of the previously fetched web page
 * @returns {string} The extracted description of the item extracted from within the provided HTML content
 */
function extractDescriptionFromThonkyResponse(htmlResponse) {
  "use strict";

  requireNonNullParameter(htmlResponse, "htmlResponse");

  var entireTag = htmlResponse.match(/(<td>Description<\/td><td><pre>[^<>)]*)(<\/pre><\/td><\/tr>)/g)[0];

  var startPos = 0 + "<td>Description</td><td><pre>".length;
  var endPos = entireTag.length - "</pre></td></tr>".length;

  var tagContents = entireTag.substring(startPos, endPos);

  return tagContents;
}

/**
 * Imports the total price of an appraisal done on Evepraisal.
 *
 * @example
 * EVEPRAISAL_TOTAL("gp5av", "buy") *
 * @param {string} appraisal_id the alphanumeric ID of the appraisal
 * @param {string} [order_type="sell"] the order type. See {@link EVEPRAISAL#ORDER_TYPE}
 * @return {number} The total value of the specified appraisal
 * @customfunction
 **/
function EVEPRAISAL_TOTAL(appraisal_id, order_type) {
  "use strict";

  requireNonNullParameter(appraisal_id, "appraisal_id");

  order_type = order_type == null ? DEFAULT.ORDER_TYPE : order_type;

  const JSON_DATA_URL = SITES.EVEPRAISAL_APPRAISAL_PREFIX + encodeURIComponent(appraisal_id) + ".json";

  var jsondata = fetchCachedUrlContent(JSON_DATA_URL, "text");
  var object = JSON.parse(jsondata);

  return object.totals[order_type];
}

/**
 * Imports the price of the item from Evepraisal.
 *
 * @param {number} item_id the numeric id of an EVE Item. See the {@link EVE_TYPENAME_TO_TYPEID} function
 * @param {string} [market="jita"] the market to price an item in. See {@link EVEPRAISAL#MARKET}
 * @param {string} [order_type="sell"] the order type. See {@link EVEPRAISAL#ORDER_TYPE}
 * @param {string} [attribute] the attribute to use when . See {@link EVEPRAISAL#ATTRIBUTE}. The default is "min" for sell and "max" for buy
 * @return {number} The numerical result of the specified attribute for the specified item at the specified market location
 * @customfunction
 **/
function EVEPRAISAL_ITEM(item_id, market, order_type, attribute) {
  "use strict";

  requireNonNullParameter(item_id, "item_id");

  market = market == null ? DEFAULT.MARKET : market;
  order_type = order_type == null ? DEFAULT.ORDER_TYPE : order_type;
  attribute = attribute == null ? (order_type == "sell" ? DEFAULT.ATTRIBUTE_SELL : DEFAULT.ATTRIBUTE_BUY) : attribute;

  const JSON_DATA_URL = SITES.EVEPRAISAL_ITEM_PREFIX + encodeURIComponent(item_id) + ".json";

  var jsondata = fetchCachedUrlContent(JSON_DATA_URL, "text");
  var object = JSON.parse(jsondata);

  for(var i in object.summaries) {
    if(object.summaries[i].market_name == market) {
      return object.summaries[i].prices[order_type][attribute];
    }
  }

  throw "market " + market + " not found";
}

/**
 * Generates an item ID from the item name by using Fuzzworks
 *
 * @param {string} item_name The item's name
 * @return {string} The item ID associated with the given item name
 * @customfunction
 */
function EVE_TYPENAME_TO_TYPEID(item_name) {
  "use strict";

  requireNonNullParameter(item_name, "item_name");

  var TYPENAME_URL = SITES.FUZZWORKS_ITEMID_PREFIX + encodeURIComponent(item_name);
  var json_data = fetchCachedUrlContent(TYPENAME_URL, "text");
  var object = JSON.parse(json_data);

  if(object.typeID != 0 && object.typeName != "bad item") {
    return object.typeID;
  } else {
    throw "item " + item_name + " not found";
  }
}

/**
 * Scrapes the given item's description from a 3rd-party site.
 *
 * @param {string\|array} item_name The item's name (or an array generated from a range)
 * @param {string} [link_type="thonky"] The site to scrape the description from (either "zKillboard","eveInfo" or "thonky")
 * @return {string\|array} The resulting description text (or an array containing the description text for the given range)
 */
function EVE_TYPENAME_TO_DESCRIPTION(item_name, link_type) {
  "use strict";

  requireNonNullParameter(item_name, "item_name");

  if(arguments[0] && arguments[0].map) {
    return multipleValueRangesHandler(EVE_TYPENAME_TO_DESCRIPTION, 2, arguments);
  }

  link_type = link_type == null ? DEFAULT.FETCHED_DESCRIPTION_TYPE : link_type;

  const type_id = EVE_TYPENAME_TO_TYPEID(item_name);

  var description_url;

  if("eveInfo".equalsIgnoreCase(link_type)) {
    description_url = SITES.EVEINFO_ITEMS_PREFIX + encodeURIComponent(type_id);
  } else if("zKillboard".equalsIgnoreCase(link_type)) {
    description_url = SITES.ZKILLBOARD_ITEMS_PREFIX + encodeURIComponent(type_id);
  } else if("thonky".equalsIgnoreCase(link_type)) {
    description_url = SITES.THONKY_ITEMS_PREFIX + encodeURIComponent(type_id);
  } else {
    throw "This error should only be reached if a new linked description was not handled in EVE_TYPENAME_TO_LINKED_DESCRIPTION";
  }

  try {
    var htmlResponse = fetchCachedUrlContent(description_url, "text");

    var tagContents;

    if("eveInfo".equalsIgnoreCase(link_type)) {
      tagContents = extractDescriptionFromEveInfoResponse(htmlResponse);
    } else if("zKillboard".equalsIgnoreCase(link_type)) {
      tagContents = extractDescriptionFromZKillboardResponse(htmlResponse);
    } else if("thonky".equalsIgnoreCase(link_type)) {
      tagContents = extractDescriptionFromThonkyResponse(htmlResponse);
    } else {
      throw "This error should only be reached if a new description source was not handled in EVE_TYPENAME_TO_DESCRIPTION";
    }

    return tagContents;
  } catch (ex) {
    throw ex;
  }
}

/**
 * Generates a URL to a webpage containing the given item's description.
 *
 * @param {string} item_name The item's name
 * @param {string} [link_type="zKillboard"] The site containing the description info which we will link to (either "zKillboard","eveInfo" or "thonky")
 * @return {string} The URL to the 3rd party item info site
 * @customfunction
 */
function EVE_TYPENAME_TO_LINKED_DESCRIPTION(item_name, link_type) {
  "use strict";

  requireNonNullParameter(item_name, "item_name");

  link_type = link_type == null || link_type == "" ? DEFAULT.LINKED_DESCRIPTION_TYPE : link_type;

  const type_id = EVE_TYPENAME_TO_TYPEID(item_name);

  var description_url;

  if("eveInfo".equalsIgnoreCase(link_type)) {
    description_url = SITES.EVEINFO_ITEMS_PREFIX + encodeURIComponent(type_id);
  } else if("zKillboard".equalsIgnoreCase(link_type)) {
    description_url = SITES.ZKILLBOARD_ITEMS_PREFIX + encodeURIComponent(type_id);
  } else if("thonky".equalsIgnoreCase(link_type)) {
    description_url = SITES.THONKY_ITEMS_PREFIX + encodeURIComponent(type_id);
  } else {
    throw "This error should only be reached if a new linked description was not handled in EVE_TYPENAME_TO_LINKED_DESCRIPTION";
  }

  return description_url;
}

/**
 * Generates a URL for an image of the given item.
 *
 * @param {string\|array} item_name The item's name
 * @param {number} [image_dimension=32] The square dimensions of the returned image (either 32 or 64)
 * @return {string} The URL to the EVE Online Image Server image representing the item
 * @customfunction
 */
function EVE_IMAGE_URL(item_name, image_dimension) {
  "use strict";

  requireNonNullParameter(item_name, "item_name");

  const numArgs = 2;
  const targetFunction = EVE_IMAGE_URL;

  if(arguments[0] && arguments[0].map) {
    return multipleValueRangesHandler(EVE_IMAGE_URL, 2, arguments);
  }

  const isImageDimensionInvalid = image_dimension == null || !(image_dimension == 32 || image_dimension == 64);

  image_dimension = isImageDimensionInvalid ? DEFAULT.IMAGE_DIMENSION : image_dimension;

  const image_server_url_prefix = SITES.EVEONLINE_IMAGESERVER_PREFIX;
  const image_server_url_suffix = "_" + encodeURIComponent(image_dimension) + ".png";

  const type_id = EVE_TYPENAME_TO_TYPEID(item_name);

  return image_server_url_prefix + encodeURIComponent(type_id) + image_server_url_suffix;
}
