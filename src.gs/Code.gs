// Compiled using ts2gas 3.6.4 (TypeScript 4.1.3)
var exports = exports || {};
var module = module || { exports: exports };
//import { stat } from "fs";
// ID of the Target Google Spreadsheet (Book)
var VAL_ID_TARGET_BOOK = '{ID of Target Google Spreadsheet}';
// ID of the Google Drive where the images will be placed
var VAL_ID_GDRIVE_FOLDER_MEDIA = '{ID of Google Drive Folder to place media files}';
// ID of the Google Drive where backup files will be placed
var VAL_ID_GDRIVE_FOLDER_BACKUP = '{ID of Google Drive Folder to place backup files}';
// ID of the Google Drive where history files will be placed
var VAL_ID_GDRIVE_FOLDER_HISTORY = '{ID of Google Drive Folder to place history files}';
// Key and Secret to access Twitter APIs
var VAL_CONSUMER_API_KEY = '{Consumer Key got from Twitter Developer Dashboard}';
var VAL_CONSUMER_API_SECRET = '{Consumer Secret value got from Twitter Developer Dashboard}';
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// DEFINES
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
var VERSION = 1.1;
var DURATION_MILLISEC_NOT_FOR_BACKUP = 7 * 24 * 60 * 60 * 1000;
var TIME_LOCALE = "JST";
var FORMAT_DATETIME_DATE = "yyyy-MM-dd";
var FORMAT_DATETIME_ISO8601_DATE = "yyyy-MM-dd";
var FORMAT_DATETIME_ISO8601_TIME = "HH:mm:ss";
var FORMAT_DATETIME_DATE_NUM = "yyyyMMdd";
var FORMAT_DATETIME = "yyyy-MM-dd (HH:mm:ss)";
var FORMAT_TIMESTAMP = "yyyyMMddHHmmss";
var NAME_SHEET_USAGE = "!USAGE";
var NAME_SHEET_LOG = "!LOG";
var NAME_SHEET_ERROR = "!ERROR";
var SHEET_NAME_COMMON_SETTINGS = "%settings";
var CELL_SETTINGS_TITLE_LIMIT = "[limitcounttweets]";
var CELL_SETTINGS_TITLE_EMAIL = "[email]";
var CELL_SETTINGS_TITLE_KEYWORD = "[Keyword]";
var CELL_SETTINGS_TITLE_BAN_WORDS = "[Banwords]";
var CELL_SETTINGS_TITLE_BAN_USERS = "[Ban Users]";
var CELL_SETTINGS_TITLE_FILTER_WORDS = "[Filter Words]";
var CELL_SETTINGS_TITLE_ALERT_THRESHOLD_FP = "[Alert Threshold Filtered Percentage]";
var CELL_SETTINGS_TITLE_DOWNLOAD_MEDIA = "[Download Media]";
var CELL_SETTINGS_TITLE_POINT_TWEET = "[Point/Tweet]";
var CELL_SETTINGS_TITLE_POINT_RETWEET = "[Point/Retweet]";
var CELL_SETTINGS_TITLE_POINT_REPLY = "[Point/Reply]";
var CELL_SETTINGS_TITLE_POINT_MEDIA = "[Point/Media]";
var CELL_SETTINGS_TITLE_POINT_FAVORITE = "[Point/Nice]";
var CELL_SETTINGS_TITLE_ALERT_THRESHOLD = "[Alert Threshold Points]";
var CELL_SETTINGS_TITLE_LAST_KEYWORD = "[[Last Keyword]]";
var CELL_SETTINGS_TITLE_LAST_UPDATED = "[[Last Updated]]";
var CELL_SETTINGS_TITLE_END_SETTINGS = "//end of settings";
var CELL_HEADER_TITLES_DATA = { tweetId: "Tweet Id", createdAt: "Created at", userId: "User Id", userName: "User name", screenName: "Screen name", retweet: "Retweet", favorite: "Favorite", tweet: "Tweet", media: "Media" };
var CELL_HEADER_TITLES_HISTORY = { dateTime: "date time", keyword: "keyword", ptTweet: "pt/Tweet", ptRetweet: "pt/Retweet", ptMedia: "pt/Media", ptFav: "pt/Fav", threshold: "Threshold", blank: "", total: "Total", tweets: "Tweets", retweets: "Retweets", media: "Media", favs: "Favs" };
var DEFAULT_LIMIT_TWEETS = 10;
var DEFAULT_DOWNLOAD_MEDIA = true;
var DEFAULT_POINT_TWEET = 10;
var DEFAULT_POINT_RETWEET = 10;
var DEFAULT_POINT_REPLY = 10;
var DEFAULT_POINT_MEDIA = 10;
var DEFAULT_POINT_FAVORITE = 10;
var DEFAULT_ALERT_THRESHOLD = 100;
var MAX_ROW_RANGE_SETTINGS = 20;
var MAX_COLUMN_RANGE_SETTINGS = 30;
var MAX_ROW_SEEK_HEADER = 20;
var DEFAULT_ALERT_THRESHOLD_RWP = 30;
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// GLOBALS
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
var g_isDebugMode = true;
var g_isEnabledLogging = true;
var g_settingsCommon = null;
var g_folderBackup = DriveApp.getFolderById(VAL_ID_GDRIVE_FOLDER_BACKUP);
var g_datetime = new Date();
var g_timestamp = TIME_LOCALE + ": " + Utilities.formatDate(g_datetime, TIME_LOCALE, FORMAT_DATETIME);
var g_folderMedia = DriveApp.getFolderById(VAL_ID_GDRIVE_FOLDER_MEDIA);
var g_folderHistory = DriveApp.getFolderById(VAL_ID_GDRIVE_FOLDER_HISTORY);
var g_book = SpreadsheetApp.openById(VAL_ID_TARGET_BOOK);
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// OBJECTS
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// Tweet
//
var Tweet = /** @class */ (function () {
    function Tweet(id_str, created_at, text, user_id_str, user_name, screen_name, list_media, retweet_count, favorite_count) {
        this.id_str = id_str;
        this.created_at = created_at;
        this.created_at_str = Utilities.formatDate(new Date(created_at), TIME_LOCALE, FORMAT_DATETIME);
        this.text = text;
        this.user_id_str = user_id_str;
        this.user_name = user_name;
        this.screen_name = screen_name;
        this.list_media = list_media;
        this.retweet_count = retweet_count;
        this.favorite_count = favorite_count;
        if (retweet_count < 0 || favorite_count < 0) {
            logOut("UNEXPECTED: id = " + id_str + ", retweet_count = " + retweet_count + ", favorite_count = " + favorite_count);
        }
    }
    return Tweet;
}());
//
// Media
//
var Media = /** @class */ (function () {
    function Media(media_url) {
        this.media_url = media_url;
    }
    return Media;
}());
//
// Settings - This object will be used for as Common settings and as Particular query settings
//
var Settings = /** @class */ (function () {
    function Settings(currentKeyword, lastKeyword, limitTweets, filterWords, fpAlertThreshold, emails, banWords, banUsers, bDownloadMedia, ptTweet, ptRetweet, ptReply, ptMedia, ptFavorite, ptAlertThreshold, rowEndSettings, headerInfo) {
        this.currentKeyword = currentKeyword;
        this.lastKeyword = lastKeyword;
        this.limitTweets = limitTweets;
        this.fpAlertThreshold = fpAlertThreshold;
        this.filterWords = filterWords;
        this.emails = emails;
        this.banWords = banWords;
        this.banUsers = banUsers;
        this.bDownloadMedia = bDownloadMedia;
        this.ptTweet = ptTweet;
        this.ptRetweet = ptRetweet;
        this.ptReply = ptReply;
        this.ptMedia = ptMedia;
        this.ptFavorite = ptFavorite;
        this.ptAlertThreshold = ptAlertThreshold;
        this.rowEndSettings = rowEndSettings;
        this.headerInfo = headerInfo;
    }
    return Settings;
}());
//
// Header
//
var HeaderInfo = /** @class */ (function () {
    function HeaderInfo() {
        this.idx_row = null;
        this.idx_tweetId = null;
        this.idx_createdAt = null;
        this.idx_userId = null;
        this.idx_userName = null;
        this.idx_screenName = null;
        this.idx_tweet = null;
        this.idx_retweet = null;
        this.idx_favorite = null;
        this.idx_media = null;
    }
    return HeaderInfo;
}());
var Stats = /** @class */ (function () {
    function Stats() {
        this.countTweets = 0;
        this.countRetweets = 0;
        this.countMedias = 0;
        this.countFavorites = 0;
        this.countFilterWords = 0;
        this.points = 0;
        this.percentWords = 0.0;
        this.isWordFilter = false;
    }
    return Stats;
}());
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//  OAuth Routines
//
// REFERENCE:
//    https://tech-cci.io/archives/4228
//    https://kazyblog.com/try-google-app-script-for-tweet
//    https://developer.twitter.com/en/docs/twitter-api/v1/tweets/search/api-reference/get-search-tweets
//
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function getOAuthURL() {
    Logger.log(getService().authorize());
}
function getService() {
    //let VAL_CONSUMER_API_KEY     = PropertiesService.getScriptProperties().getProperty("CONSUMER_API_KEY");
    //let VAL_CONSUMER_API_SECRET  = PropertiesService.getScriptProperties().getProperty("CONSUMER_API_SECRET");
    return (OAuth1.createService("Twitter")
        .setAccessTokenUrl("https://api.twitter.com/oauth/access_token")
        .setRequestTokenUrl("https://api.twitter.com/oauth/request_token")
        .setAuthorizationUrl("https://api.twitter.com/oauth/authorize")
        .setConsumerKey(VAL_CONSUMER_API_KEY)
        .setConsumerSecret(VAL_CONSUMER_API_SECRET)
        .setCallbackFunction("authCallback")
        .setPropertyStore(PropertiesService.getUserProperties()));
}
function resetService() {
    var service = getService();
    service.reset();
}
function authCallback(request) {
    var service = getService();
    var authorized = service.handleCallback(request);
    if (authorized) {
        return HtmlService.createHtmlOutput("success!!");
    }
    else {
        return HtmlService.createHtmlOutput("failed");
    }
}
// ====================================================================================================================
//
// Debug / Log
//
// ====================================================================================================================
//
// Name: dbgOut
// Desc:
//
function dbgOut(text) {
    if (!g_isDebugMode) {
        return;
    }
    Logger.log(text);
}
//
// Name: logOut
// Desc:
//
function logOut(text) {
    text = g_timestamp + "\t" + text;
    if (!g_isEnabledLogging) {
        return;
    }
    gsAddLineAtBottom(NAME_SHEET_LOG, text);
}
//
// Name: errOut
// Desc:
//
function errOut(text) {
    text = g_timestamp + "\t" + text;
    gsAddLineAtBottom(NAME_SHEET_ERROR, text);
}
// ====================================================================================================================
//
// Basic Twitter Functions
//
// ====================================================================================================================
//
// Func: twGetTimeLine
// Desc:
//    sample func to see how Twitter API works
// Return:
//    none
//
function twGetTimeLine() {
    var twitterService = getService();
    if (twitterService.hasAccess()) {
        var twMethod = { method: "GET" };
        var json = twitterService.fetch("https://api.twitter.com/1.1/statuses/home_timeline.json?count=5", twMethod);
        var array = JSON.parse(json);
        var tweets = array.map(function (tweet) {
            return tweet.text;
        });
        Logger.log(tweets);
    }
    else {
        Logger.log(service.getLastError());
    }
}
//
// Func: twSearchTweet
// Desc:
//   search tweet based on the specified keyword
// Return:
//   'Tweet' object
//
function twSearchTweet(keywords, maxCount) {
    if (maxCount === void 0) { maxCount = DEFAULT_LIMIT_TWEETS; }
    var encodedKeyword = keywords.trim().replace(/\s+/g, '+AND+');
    encodedKeyword = encodeURIComponent(encodedKeyword);
    try {
        var twitterService = getService();
        if (!twitterService.hasAccess()) {
            errOut(twitterService.getLastError());
            return null;
        }
        var url = 'https://api.twitter.com/1.1/search/tweets.json?q='
            + encodedKeyword
            + '&result_type=recent&lang=ja&locale=ja&count='
            + maxCount;
        var response = twitterService.fetch(url, { method: "GET" });
        var tweets = [];
        var json = JSON.parse(response);
        json.statuses.forEach(function (tweet) {
            if (!tweet.retweeted_status) {
                // INFO: How to get images via Twitter API
                // https://qiita.com/w_cota/items/a87b421ba8bc2b90a938
                var list_media = [];
                if (tweet.entities.media != undefined && tweet.entities.media[0].type == 'photo') {
                    for (var i = 0; i < tweet.extended_entities.media.length; i++) {
                        var media_url = tweet.extended_entities.media[i].media_url;
                        list_media.push(new Media(media_url));
                    }
                }
                tweets.push(new Tweet(tweet.id_str, tweet.created_at, tweet.text, tweet.user.id_str, tweet.user.name, tweet.user.screen_name, list_media, tweet.retweet_count, tweet.favorite_count));
            }
        });
        return tweets;
    }
    catch (ex) {
        errOut(ex);
        errOut("keyword [" + keywords + "] encoded keyword = [" + encodedKeyword + "]");
        return null;
    }
}
// ====================================================================================================================
//
// Utilities for Misc
//
// ====================================================================================================================
//
// Name: utIsValid
// Desc:
//
function utIsValid(val) {
    return val == null || val == undefined;
}
//
// Name: utGetDateTime
// Desc:
//
function utGetDateTime() {
    return (TIME_LOCALE +
        ": " +
        Utilities.formatDate(new Date(), TIME_LOCALE, FORMAT_DATETIME));
}
// ====================================================================================================================
//
// Utilities for Google Spreadsheet
//
// ====================================================================================================================
//
// Name: gsAddLineAtLast
// Desc:
//  Add the specified text at the bootom of the specified sheet.
//
function gsAddLineAtBottom(sheetName, text) {
    try {
        var sheet = g_book.getSheetByName(sheetName);
        if (!sheet) {
            sheet = g_book.insertSheet(sheetName, g_book.getNumSheets());
        }
        var lastRow = sheet.getLastRow();
        if (lastRow == sheet.getMaxRows()) {
            sheet.insertRowsAfter(lastRow, 1);
        }
        var range = sheet.getRange(lastRow + 1, 1, 1, 2);
        if (range) {
            var valsRange = range.getValues();
            var row = valsRange[0];
            row[0] = g_timestamp;
            row[1] = String(text);
            range.setValues(valsRange);
        }
    }
    catch (e) {
        Logger.log("EXCEPTION: gsAddLineAtBottom: " + e.message);
    }
}
//
// Name: gsClearData
// Desc:
//  Remove all rows from ROW_DATA_START.
//
function gsClearData(sheet, rowDataStart) {
    var lastRow = sheet.getRange("A:A").getLastRow();
    if (rowDataStart - 1 < lastRow) {
        sheet.deleteRows(rowDataStart, lastRow - rowDataStart + 1);
    }
}
// ====================================================================================================================
//
// Google Spreadsheet Routines
//
// ====================================================================================================================
//
// Name: sendMail
// Desc:
//
//
function sendMail(sheet, stats, settings) {
    var subject = "";
    if (stats.isWordFilter) {
        subject += "★";
    }
    subject += "⚡AutoMail:TweetWatcher⚡: [" + sheet.getName() + "]:[" + settings.currentKeyword + "]: pt=" + stats.points;
    if (stats.isWordFilter) {
        subject += ", %=" + stats.percentWords;
    }
    var body = "Total Points : " + stats.points + ", " + "Perentage of Filtered Words: " + stats.percentWords;
    var htmlBody = "Total Points : " + getPoint(stats, settings) + "<br>";
    if (stats.isWordFilter) {
        htmlBody += "★Percentage of Filtered Words : " + stats.percentWords + "% (Threshold=" + settings.fpAlertThreshold + "%)<br>";
    }
    htmlBody += "Threshold (point): " + settings.ptAlertThreshold + "<br>";
    if (stats.isWordFilter) {
        htmlBody += "★Threshold (percent): " + settings.fpAlertThreshold + "<br>";
    }
    htmlBody += ("<br>"
        + "<a href='https://docs.google.com/spreadsheets/d/" + VAL_ID_TARGET_BOOK + "/edit'>&lt;Link to the Spreadsheet&gt;</a><br>"
        + "Sheet name: " + sheet.getName() + "<br>"
        + "Keyword: " + settings.currentKeyword + "<br>"
        + "<br>"
        + "Stats:<br>"
        + "<table>"
        + "<tr><th>Item</th><th>Count</th><th>Pt/Item</th><th>Points</th></tr>"
        + "<tr><td>Tweet</td><td>" + stats.countTweets + "</td><td>" + settings.ptTweet + "</td><td>" + (stats.countTweets * settings.ptTweet) + "</td></tr>"
        + "<tr><td>Retweet</td><td>" + stats.countRetweets + "</td><td>" + settings.ptRetweet + "</td><td>" + (stats.countRetweets * settings.ptRetweet) + "</td></tr>"
        + "<tr><td>Favorite</td><td>" + stats.countFavorites + "</td><td>" + settings.ptFavorite + "</td><td>" + (stats.countFavorites * settings.ptFavorite) + "</td></tr>"
        + "<tr><td>Media</td><td>" + stats.countMedias + "</td><td>" + settings.ptMedia + "</td><td>" + (stats.countMedias * settings.ptMedia) + "</td></tr>"
        + "</table>");
    settings.emails.forEach(function (emailAddr) {
        GmailApp.sendEmail(emailAddr, subject, body, { htmlBody: htmlBody });
    });
}
//
// Name: getPoint
// Desc:
//
//
function getPoint(stats, settings) {
    return settings.ptTweet * stats.countTweets
        + settings.ptRetweet * stats.countRetweets
        + settings.ptMedia * stats.countMedias
        + settings.ptFavorite * stats.countFavorites;
}
//
// Name: insertHeaderData
// Desc:
//  Insert the header row at the specified row (1-based)
//
function inseartHeaderData(sheet, row) {
    var rangeHeader = sheet.getRange(row, 1, 1, Object.keys(CELL_HEADER_TITLES_DATA).length);
    var valsHeader = rangeHeader.getValues();
    Object.keys(CELL_HEADER_TITLES_DATA).forEach(function (val, c) {
        valsHeader[0][c] = val;
    });
    rangeHeader.setValues(valsHeader);
}
//
// Name: insertHeaderHistory
// Desc:
//  Insert the header row at the specified row (1-based)
//
function inseartHeaderHistory(sheet, row) {
    var rangeHeader = sheet.getRange(row, 1, 1, Object.keys(CELL_HEADER_TITLES_HISTORY).length);
    var valsHeader = rangeHeader.getValues();
    Object.keys(CELL_HEADER_TITLES_HISTORY).forEach(function (val, c) {
        valsHeader[0][c] = val;
    });
    rangeHeader.setValues(valsHeader);
}
//
// Name: getSheet
// Desc:
//  Get access to the specified sheet in the specified book, and if the book and the sheet don't exist, create them accordingly.
//  This function is used for backup feature and storing history data.
//
function getSheet(folderParent, pathTarget, nameSheet, indexSheet, drawHeaderFunc) {
    var pathSplit = pathTarget.split("/");
    var listPath = [];
    var nameFile = null;
    for (var i = 0; i < pathSplit.length; i++) {
        if (pathSplit[i].length > 0) {
            if (i == pathSplit.length - 1) {
                nameFile = pathSplit[i];
                break;
            }
            listPath.push(pathSplit[i]);
        }
    }
    if (null == nameFile) {
        errOut("getFile() - wrong path name - " + pathTarget);
        return null;
    }
    var folderTarget = folderParent;
    for (var i = 0; i < listPath.length; i++) {
        if (folderTarget.getFoldersByName(listPath[i]).hasNext()) {
            folderTarget = folderTarget.getFoldersByName(listPath[i]).next();
        }
        else {
            folderTarget = folderTarget.createFolder(listPath[i]);
        }
    }
    var fileTarget = folderTarget.getFilesByName(nameFile);
    var book = null;
    var sheet = null;
    if (fileTarget && fileTarget.hasNext()) {
        var file = fileTarget.next();
        book = SpreadsheetApp.openById(file.getId());
        sheet = book.getSheetByName(nameSheet);
        if (!sheet) {
            sheet = book.insertSheet(nameSheet, indexSheet);
            drawHeaderFunc(sheet, 1);
        }
    }
    else {
        book = SpreadsheetApp.create(nameFile);
        book.getActiveSheet().setName(nameSheet);
        var idBook = DriveApp.getFileById(book.getId());
        folderTarget.addFile(idBook);
        sheet = book.getActiveSheet();
        drawHeaderFunc(sheet, 1);
    }
    return sheet;
}
//
// Name: addHistory
// Desc:
//
//
function addHistory(nameSheet, indexSheet, stats, settings) {
    var fullYear = g_datetime.getFullYear();
    var month = 1 + g_datetime.getMonth();
    var nameBook = g_book.getName() + "/" + String(fullYear) + "_" + ('00' + month).slice(-2);
    var sheetHistory = getSheet(g_folderHistory, nameBook, nameSheet, indexSheet, inseartHeaderHistory);
    var lastRow = sheetHistory.getLastRow();
    var rowToWrite = lastRow + 1;
    if (rowToWrite > sheetHistory.getMaxRows()) {
        sheetHistory.insertRowsAfter(lastRow, 1);
    }
    var rangeLastRow = sheetHistory.getRange(rowToWrite, 1, 1, Object.keys(CELL_HEADER_TITLES_HISTORY).length);
    if (rangeLastRow) {
        var valsRange = rangeLastRow.getValues();
        valsRange[0][0] = g_timestamp;
        valsRange[0][1] = settings.currentKeyword;
        valsRange[0][2] = settings.ptTweet;
        valsRange[0][3] = settings.ptRetweet;
        valsRange[0][4] = settings.ptMedia;
        valsRange[0][5] = settings.ptFavorite;
        valsRange[0][6] = settings.ptAlertThreshold;
        valsRange[0][7] = "";
        valsRange[0][8] = getPoint(stats, settings);
        ;
        valsRange[0][9] = stats.countTweets;
        valsRange[0][10] = stats.countRetweets;
        valsRange[0][11] = stats.countMedias;
        valsRange[0][12] = stats.countFavorites;
        rangeLastRow.setValues(valsRange);
    }
}
//
// Name: downloadMedia
// Desc:
//  Download media used in a tweet in the date folder.
//
function downloadMedia(tweet, dateCreatedAt, settings) {
    var folderMedia = null;
    if (0 < tweet.list_media.length && settings.bDownloadMedia) {
        var strDate = Utilities.formatDate(dateCreatedAt, TIME_LOCALE, FORMAT_DATETIME_DATE);
        var foldersOfDate = g_folderMedia.getFoldersByName(strDate);
        var folderDate = null;
        if (foldersOfDate.hasNext()) {
            folderDate = foldersOfDate.next();
        }
        else {
            folderDate = g_folderMedia.createFolder(strDate);
        }
        folderMedia = folderDate.createFolder(tweet.id_str);
        tweet.list_media.forEach(function (media) {
            // console.log( media );
            var imageBlob = UrlFetchApp.fetch(media.media_url).getBlob();
            folderMedia.createFile(imageBlob);
        });
    }
    return folderMedia;
}
//
// Name: addTweetDataAtBottom
// Desc:
//  Add the specified text at the bootom of the specified sheet.
//
function addTweetDataAtBottom(sheet, indexSheet, tweet, settings, stats) {
    try {
        var lastRow = sheet.getLastRow();
        var rowToWrite = lastRow + 1;
        if (rowToWrite > sheet.getMaxRows()) {
            sheet.insertRowsAfter(lastRow, 1);
        }
        var dateCreatedAt = new Date(tweet.created_at);
        var folderMedia = downloadMedia(tweet, dateCreatedAt, settings);
        var rangeLastRow = sheet.getRange(rowToWrite, 1, 1, Object.keys(CELL_HEADER_TITLES_DATA).length);
        if (rangeLastRow) {
            var valsRange = rangeLastRow.getValues();
            var row = valsRange[0];
            row[settings.headerInfo.idx_tweetId] = '=HYPERLINK("https://twitter.com/' + tweet.screen_name + '/status/' + tweet.id_str + '", "' + tweet.id_str + '")';
            row[settings.headerInfo.idx_createdAt] = tweet.created_at_str;
            row[settings.headerInfo.idx_userId] = tweet.user_id_str;
            row[settings.headerInfo.idx_userName] = tweet.user_name;
            row[settings.headerInfo.idx_screenName] = tweet.screen_name;
            row[settings.headerInfo.idx_retweet] = tweet.retweet_count;
            row[settings.headerInfo.idx_favorite] = tweet.favorite_count;
            row[settings.headerInfo.idx_tweet] = tweet.text;
            if (folderMedia) {
                row[settings.headerInfo.idx_media] = folderMedia.getUrl();
            }
            rangeLastRow.setValues(valsRange);
        }
    }
    catch (e) {
        Logger.log("EXCEPTION: addTweetDataAtBottom: " + e.message);
    }
}
//
// Name: updateExistingTweet
// Desc:
//  Check if the specified tweet contains any ban words in the specified array of the ban words.
//
function updateExistingTweet(range, valsRange, idxRow, tweet, settings, stats) {
    try {
        var prevCountRetweets = valsRange[idxRow][settings.headerInfo.idx_retweet];
        var prevCountFavorites = valsRange[idxRow][settings.headerInfo.idx_favorite];
        stats.countRetweets += (tweet.retweet_count - prevCountRetweets);
        stats.countFavorites += (tweet.favorite_count - prevCountFavorites);
        valsRange[idxRow][settings.headerInfo.idx_retweet] = tweet.retweet_count;
        valsRange[idxRow][settings.headerInfo.idx_favorite] = tweet.favorite_count;
        range.setValues(valsRange); // Update the sheet
        // logOut( "Updated: idx row = [" + idxRow + "], id=[" + tweet.id_str + "], prev # of retweets = " + prevCountRetweets + ", new # of retweets = " + tweet.retweet_count );
    }
    catch (e) {
        Logger.log("EXCEPTION: aupdateExistingTweet: " + e.message);
    }
}
//
// Name: hasWords
// Desc:
//  Check if the specified tweet contains any words in the specified array.
//
function hasWords(tweet, words) {
    for (var i = 0; i < words.length; i++) {
        //let regex = new RegExp(banWords[i]);
        //if ( -1 != tweet.text.search(regex) ) {
        if (tweet.text.includes(words[i])) {
            return true;
        }
    }
    return false;
}
//
// Name: isFromBanUsers
// Desc:
//  Check if the specified tweet contains any ban words in the specified array of the ban words.
//
function isFromBanUsers(tweet, banUsers) {
    for (var i = 0; i < banUsers.length; i++) {
        //if ( tweet.screen_name == banUsers[i] ) {
        if (tweet.user_id_str == banUsers[i]) {
            return true;
        }
    }
    return false;
}
//
// Name: isSafeTweet
// Desc:
//
function isSafeTweet(tweet, settings) {
    if (hasWords(tweet, settings.banWords) || isFromBanUsers(tweet, settings.banUsers)) {
        return false;
    }
    return true;
}
//
// Name: isValidTweet
// Desc:
//  Check if the specified tweet has all keywords
//
function isValidTweet(tweet, settings) {
    var keywords = settings.currentKeyword.split(/\s/);
    for (var i = 0; i < keywords.length; i++) {
        if ((keywords[i]).length > 0) {
            if (!tweet.text.includes(keywords[i])) {
                return false;
            }
        }
    }
    return true;
}
//
// Name: getRowIndexTweets
// Desc:
//  Check if the specified Tweet has already been recorded (true) or not (false).
//
function getRowIndexTweets(valsRange, tweet, settings) {
    var r = 0;
    for (; r < valsRange.length; r++) {
        var row = valsRange[r];
        if (!row[settings.headerInfo.idx_tweetId]) {
            break;
        }
        if (row[settings.headerInfo.idx_tweetId] == tweet.id_str) {
            return r;
        }
    }
    return -1;
}
//
// Name: isNeededToBeAdded
// Desc:
//
function isNeededToBeAdded(sheet, tweet, settings) {
    var lastRow = sheet.getLastRow();
    if (0 < lastRow - (settings.headerInfo.idx_row + 1)) {
        var range = sheet.getRange(settings.headerInfo.idx_row + 2, 1, lastRow - (settings.headerInfo.idx_row + 1), Object.keys(CELL_HEADER_TITLES_DATA).length);
        if (range) {
            var valsRange = range.getValues();
            if (-1 == getRowIndexTweets(valsRange, tweet, settings)) {
                return false;
            }
        }
    }
    return true;
}
//
// Name: updateSheet
// Desc:
//  Add a new row and write the new Tweet data.
// Return:
//  Total point per update
//
function updateSheet(sheet, indexSheet, settings, tweets) {
    var stats = new Stats();
    var lastRow = sheet.getLastRow();
    var isInitial = true;
    var range = null;
    var valsRange = null;
    if (0 < lastRow - (settings.headerInfo.idx_row + 1)) {
        isInitial = false;
        range = sheet.getRange(settings.headerInfo.idx_row + 2, 1, lastRow - (settings.headerInfo.idx_row + 1), Object.keys(CELL_HEADER_TITLES_DATA).length);
        if (range) {
            valsRange = range.getValues();
        }
    }
    for (var i = tweets.length - 1; i >= 0; i--) {
        if (!isSafeTweet(tweets[i], settings)) {
            continue;
        }
        if (!isValidTweet(tweets[i], settings)) {
            continue;
        }
        stats.countTweets++;
        stats.countMedias += (tweets[i].list_media.length > 0) ? 1 : 0;
        stats.countRetweets += tweets[i].retweet_count;
        stats.countFavorites += tweets[i].favorite_count;
        if (0 < settings.filterWords.length && hasWords(tweets[i], settings.filterWords)) {
            stats.countFilterWords++;
        }
        if (isInitial) {
            // pure new Tweet which needs to be ADDED
            addTweetDataAtBottom(sheet, indexSheet, tweets[i], settings, stats);
        }
        else if (range && valsRange) {
            addTweetDataAtBottom(sheet, indexSheet, tweets[i], settings, stats);
            /*
            let idxRow = getRowIndexTweets(valsRange, tweets[i], settings);
            if ( -1 == idxRow ) {
              // pure new Tweet which needs to be ADDED
              addTweetDataAtBottom(sheet, indexSheet, tweets[i], settings, stats);
            } else {
              // already recorded Tweet which needs to be UPDATED
              updateExistingTweet(range, valsRange, idxRow, tweets[i], settings, stats);
            }
            */
        }
    }
    return stats;
}
//
// Name: getSettings
// Desc:
//  Get settings from the common setting sheet and each query sheet.
//  This function supports both the Common Setting Sheet and settings for each query.
//
function getSettings(sheet) {
    var range = sheet.getRange(1, 1, MAX_ROW_RANGE_SETTINGS, MAX_COLUMN_RANGE_SETTINGS);
    if (range == null) {
        return;
    }
    var currentKeyword = "";
    var lastKeyword = "";
    var limitTweets = null;
    var filterWords = [];
    var fpAlertThreshold = null;
    var emails = [];
    var banWords = [];
    var banUsers = [];
    var bDownloadMedia = null;
    var ptTweet = null;
    var ptRetweet = null;
    var ptReply = null;
    var ptMedia = null;
    var ptFavorite = null;
    var ptAlertThreshold = null;
    var rowEndSettings = null;
    var valsRange = range.getValues();
    for (var r = 0; r < valsRange.length; r++) {
        var row = valsRange[r];
        var title = String(row[0]);
        if (title) {
            title = title.replace(/\s+/g, '').toLowerCase().trim();
            switch (title) {
                case CELL_SETTINGS_TITLE_LIMIT.replace(/\s+/g, '').toLowerCase().trim():
                    if (row[1]) {
                        limitTweets = Number(row[1]);
                    }
                    break;
                case CELL_SETTINGS_TITLE_KEYWORD.replace(/\s+/g, '').toLowerCase().trim():
                    if (row[1]) {
                        currentKeyword = String(row[1]).trim();
                        row[1] = currentKeyword;
                    }
                    break;
                case CELL_SETTINGS_TITLE_EMAIL.replace(/\s+/g, '').toLowerCase().trim():
                    for (var c = 1; c < row.length; c++) {
                        var val = String(row[c]).trim();
                        if (!val)
                            break;
                        emails.push(val);
                    }
                    break;
                case CELL_SETTINGS_TITLE_BAN_WORDS.replace(/\s+/g, '').toLowerCase().trim():
                    for (var c = 1; c < row.length; c++) {
                        var val = String(row[c]).trim();
                        if (!val)
                            break;
                        banWords.push(val);
                    }
                    break;
                case CELL_SETTINGS_TITLE_BAN_USERS.replace(/\s+/g, '').toLowerCase().trim():
                    for (var c = 1; c < row.length; c++) {
                        var val = String(row[c]).trim();
                        if (!val)
                            break;
                        banUsers.push(val);
                    }
                    break;
                case CELL_SETTINGS_TITLE_DOWNLOAD_MEDIA.replace(/\s+/g, '').toLowerCase().trim():
                    {
                        var val = String(row[1]).trim();
                        if (!val)
                            break;
                        bDownloadMedia = (val.toLowerCase() == "yes");
                    }
                    break;
                case CELL_SETTINGS_TITLE_POINT_TWEET.replace(/\s+/g, '').toLowerCase().trim():
                    {
                        var val = String(row[1]).trim();
                        if (!val)
                            break;
                        ptTweet = Number(val);
                    }
                    break;
                case CELL_SETTINGS_TITLE_POINT_RETWEET.replace(/\s+/g, '').toLowerCase().trim():
                    {
                        var val = String(row[1]).trim();
                        if (!val)
                            break;
                        ptRetweet = Number(val);
                    }
                    break;
                case CELL_SETTINGS_TITLE_POINT_TWEET.replace(/\s+/g, '').toLowerCase().trim():
                    {
                        var val = String(row[1]).trim();
                        if (!val)
                            break;
                        ptReply = Number(val);
                    }
                    break;
                case CELL_SETTINGS_TITLE_POINT_MEDIA.replace(/\s+/g, '').toLowerCase().trim():
                    {
                        var val = String(row[1]).trim();
                        if (!val)
                            break;
                        ptMedia = Number(val);
                    }
                    break;
                case CELL_SETTINGS_TITLE_POINT_FAVORITE.replace(/\s+/g, '').toLowerCase().trim():
                    {
                        var val = String(row[1]).trim();
                        if (!val)
                            break;
                        ptFavorite = Number(val);
                    }
                    break;
                case CELL_SETTINGS_TITLE_ALERT_THRESHOLD.replace(/\s+/g, '').toLowerCase().trim():
                    {
                        var val = String(row[1]).trim();
                        if (!val)
                            break;
                        ptAlertThreshold = Number(val);
                    }
                    break;
                case CELL_SETTINGS_TITLE_LAST_KEYWORD.replace(/\s+/g, '').toLowerCase().trim():
                    {
                        var val = String(row[1]);
                        if (val) {
                            lastKeyword = val;
                        }
                        row[1] = currentKeyword;
                    }
                    break;
                case CELL_SETTINGS_TITLE_LAST_UPDATED.replace(/\s+/g, '').toLowerCase().trim():
                    {
                        row[1] = g_timestamp;
                    }
                    break;
                case CELL_SETTINGS_TITLE_END_SETTINGS.replace(/\s+/g, '').toLowerCase().trim():
                    {
                        rowEndSettings = r;
                    }
                    break;
                case CELL_SETTINGS_TITLE_FILTER_WORDS.replace(/\s+/g, '').toLowerCase().trim():
                    for (var c = 1; c < row.length; c++) {
                        var val = String(row[c]).trim();
                        if (!val)
                            break;
                        filterWords.push(val);
                    }
                    break;
                case CELL_SETTINGS_TITLE_ALERT_THRESHOLD_FP.replace(/\s+/g, '').toLowerCase().trim():
                    {
                        var val = String(row[1]).trim();
                        if (!val)
                            break;
                        fpAlertThreshold = Number(val);
                    }
                    break;
            }
            if (rowEndSettings) {
                break;
            }
        }
    }
    range.setValues(valsRange);
    return new Settings(currentKeyword, lastKeyword, limitTweets, filterWords, fpAlertThreshold, emails, banWords, banUsers, bDownloadMedia, ptTweet, ptRetweet, ptReply, ptMedia, ptFavorite, ptAlertThreshold, rowEndSettings, null);
}
//
// Name: getSettingsActual
// Desc:
//  Returns the actual settings which cover the common settings and local settings for each sheet.
//
function getSettingsActual(settingsCommon, settingsLocal) {
    return new Settings(settingsLocal.currentKeyword, settingsLocal.lastKeyword, settingsLocal.limitTweets ? settingsLocal.limitTweets : (settingsCommon.limitTweets ? settingsCommon.limitTweets : DEFAULT_LIMIT_TWEETS), settingsCommon.filterWords.concat(settingsLocal.filterWords), settingsLocal.fpAlertThreshold ? settingsLocal.fpAlertThreshold : (settingsCommon.fpAlertThreshold ? settingsCommon.fpAlertThreshold : DEFAULT_ALERT_THRESHOLD_RWP), settingsCommon.emails.concat(settingsLocal.emails), settingsCommon.banWords.concat(settingsLocal.banWords), settingsCommon.banUsers.concat(settingsLocal.banUsers), settingsCommon.bDownloadMedia ? settingsCommon.bDownloadMedia : (settingsLocal.bDownloadMedia ? settingsLocal.bDownloadMedia : DEFAULT_DOWNLOAD_MEDIA), settingsLocal.ptTweet ? settingsLocal.ptTweet : (settingsCommon.ptTweet ? settingsCommon.ptTweet : DEFAULT_POINT_TWEET), settingsLocal.ptRetweet ? settingsLocal.ptRetweet : (settingsCommon.ptRetweet ? settingsCommon.ptRetweet : DEFAULT_POINT_RETWEET), settingsLocal.ptReply ? settingsLocal.ptReply : (settingsCommon.ptReply ? settingsCommon.ptReply : DEFAULT_POINT_REPLY), settingsLocal.ptMedia ? settingsLocal.ptMedia : (settingsCommon.ptMedia ? settingsCommon.ptMedia : DEFAULT_POINT_MEDIA), settingsLocal.ptFavorite ? settingsLocal.ptFavorite : (settingsCommon.ptFavorite ? settingsCommon.ptFavorite : DEFAULT_POINT_FAVORITE), settingsLocal.ptAlertThreshold ? settingsLocal.ptAlertThreshold : (settingsCommon.ptAlertThreshold ? settingsCommon.ptAlertThreshold : DEFAULT_ALERT_THRESHOLD), settingsLocal.rowEndSettings, null);
}
//
// Name: getHeaders
// Desc:
//
//
function getHeaderInfo(sheet) {
    var range = sheet.getRange(1, 1, MAX_ROW_SEEK_HEADER, Object.keys(CELL_HEADER_TITLES_DATA).length);
    if (range == null) {
        return;
    }
    var valsRange = range.getValues();
    var r = 0, c;
    var row;
    var rowHeader = -1;
    for (; r < valsRange.length; r++) {
        row = valsRange[r];
        for (c = 0; c < row.length; c++) {
            var title = String(row[c]);
            title = title.replace(/\s+/g, '').toLowerCase().trim();
            if (title == CELL_HEADER_TITLES_DATA.tweetId.replace(/\s+/g, '').toLowerCase().trim()) {
                rowHeader = r;
                break;
            }
        }
        if (rowHeader >= 0) {
            break;
        }
    }
    if (rowHeader == -1) {
        return null;
    }
    var headerInfo = new HeaderInfo();
    headerInfo.idx_row = r;
    for (c = 0; c < valsRange.length; c++) {
        var title = String(row[c]);
        title = title.replace(/\s+/g, '').toLowerCase().trim();
        switch (title) {
            case "tweetid":
                headerInfo.idx_tweetId = c;
                break;
            case "createdat":
                headerInfo.idx_createdAt = c;
                break;
            case "userid":
                headerInfo.idx_userId = c;
                break;
            case "username":
                headerInfo.idx_userName = c;
                break;
            case "screenname":
                headerInfo.idx_screenName = c;
                break;
            case "retweet":
                headerInfo.idx_retweet = c;
                break;
            case "favorite":
                headerInfo.idx_favorite = c;
                break;
            case "tweet":
                headerInfo.idx_tweet = c;
                break;
            case "media":
                headerInfo.idx_media = c;
                break;
        }
    }
    if (null == headerInfo.idx_row
        || null == headerInfo.idx_tweetId
        || null == headerInfo.idx_createdAt
        || null == headerInfo.idx_userId
        || null == headerInfo.idx_userName
        || null == headerInfo.idx_screenName
        || null == headerInfo.idx_tweet
        || null == headerInfo.idx_retweet
        || null == headerInfo.idx_favorite
        || null == headerInfo.idx_media) {
        return null;
    }
    return headerInfo;
}
//
// Name: main
// Desc:
//  Entry point of this program.
//
function main() {
    var sheets = g_book.getSheets();
    var indexSheet = 0;
    sheets.forEach(function (sheet) {
        var sheetName = sheet.getName();
        //if (sheetName.toLocaleLowerCase().trim().startsWith('!')) {
        if (sheetName.match(/^\!.*/)) {
            return;
        }
        //
        // loading common settings
        //
        if (sheetName.toLocaleLowerCase().trim() === SHEET_NAME_COMMON_SETTINGS) {
            g_settingsCommon = getSettings(sheet);
            return;
        }
        //
        // loading local settings
        //
        var settingsLocal = getSettings(sheet);
        if (!settingsLocal.currentKeyword) {
            return;
        }
        //
        // create temporary settings which cover common and local settings
        //
        var settingsActual = getSettingsActual(g_settingsCommon, settingsLocal);
        //
        // get the header row info to add data
        //
        var headerInfo = getHeaderInfo(sheet);
        if (!headerInfo) {
            return;
        }
        settingsActual.headerInfo = headerInfo;
        if (!settingsLocal.currentKeyword || settingsLocal.lastKeyword != settingsLocal.currentKeyword) {
            gsClearData(sheet, headerInfo.idx_row + 2);
        }
        var tweets = twSearchTweet(settingsActual.currentKeyword);
        if (null != tweets && 0 < tweets.length) {
            var stats = updateSheet(sheet, indexSheet, settingsActual, tweets);
            stats.points = getPoint(stats, settingsActual);
            stats.percentWords = 100.0 * (stats.countFilterWords) / stats.countTweets;
            stats.isWordFilter = settingsActual.fpAlertThreshold < stats.percentWords;
            if (stats.points > settingsActual.ptAlertThreshold || stats.isWordFilter) {
                sendMail(sheet, stats, settingsActual);
            }
            addHistory(sheet.getName(), indexSheet, stats, settingsActual);
        }
        indexSheet++;
    });
}
//
// Name: duplicateBook
// Desc:
//  Duplicate the specified book at the specified drive
//
function duplicateBook(dateNow) {
    try {
        // crete a backup folder of the day
        g_folderBackup = DriveApp.getFolderById(VAL_ID_GDRIVE_FOLDER_BACKUP);
        var strDate = Utilities.formatDate(dateNow, TIME_LOCALE, FORMAT_DATETIME_DATE);
        var foldersOfDate = g_folderBackup.getFoldersByName(strDate);
        var folderDate = null;
        if (foldersOfDate.hasNext()) {
            folderDate = foldersOfDate.next();
        }
        else {
            folderDate = g_folderBackup.createFolder(strDate);
        }
        // duplicate the working spreadsheet
        var strDateISO8601 = Utilities.formatDate(dateNow, TIME_LOCALE, FORMAT_DATETIME_ISO8601_DATE) + "T" + Utilities.formatDate(dateNow, TIME_LOCALE, FORMAT_DATETIME_ISO8601_TIME) + "+09:00";
        var fileTarget = DriveApp.getFileById(VAL_ID_TARGET_BOOK);
        var nameFileBackup = "BACKUP_" + strDateISO8601 + "_" + fileTarget.getName();
        fileTarget.makeCopy(nameFileBackup, folderDate);
        return true;
    }
    catch (ex) {
        errOut(ex);
        return false;
    }
}
//
// Name: moveData
// Desc:
//
//
function moveData(nameBookSrc, fullYear, month, nameSheet, indexSheet, valsRangeSrc, rowStart, rowNum) {
    try {
        var nameBook = nameBookSrc + "/" + String(fullYear) + "_" + ('00' + month).slice(-2);
        var sheetBackup = getSheet(g_folderBackup, nameBook, nameSheet, indexSheet, inseartHeaderData);
        var lastRowBackup = sheetBackup.getLastRow();
        var maxRowsBackup = sheetBackup.getMaxRows();
        if (maxRowsBackup - lastRowBackup < rowNum) {
            sheetBackup.insertRowsAfter(sheetBackup.getMaxRows(), rowNum - (maxRowsBackup - lastRowBackup));
        }
        var rangeBackup = sheetBackup.getRange(sheetBackup.getLastRow() + 1, 1, rowNum, Object.keys(CELL_HEADER_TITLES_DATA).length);
        if (rangeBackup) {
            var valsRangeBackup = rangeBackup.getValues();
            for (var r = 0; r < rowNum; r++) {
                for (var c = 0; c < Object.keys(CELL_HEADER_TITLES_DATA).length; c++) {
                    valsRangeBackup[r][c] = valsRangeSrc[rowStart + r][c];
                }
                valsRangeBackup[r][0] = '=HYPERLINK("https://twitter.com/' + valsRangeBackup[r][4] + '/status/' + valsRangeBackup[r][0] + '", "' + valsRangeBackup[r][0] + '")';
            }
            rangeBackup.setValues(valsRangeBackup);
        }
        return true;
    }
    catch (ex) {
        errOut("moveData: " + ex);
        return false;
    }
}
//
// Name: backup
// Desc:
//  Entry point for backup
//
function backup() {
    var sheets = g_book.getSheets();
    var nameBook = g_book.getName();
    var indexSheet = 0;
    sheets.forEach(function (sheet) {
        var sheetName = sheet.getName();
        if (sheetName.toLocaleLowerCase().trim().startsWith('!')) {
            return;
        }
        if (sheetName.toLocaleLowerCase().trim() === SHEET_NAME_COMMON_SETTINGS) {
            return;
        }
        var headerInfo = getHeaderInfo(sheet);
        if (!headerInfo) {
            return;
        }
        var lastFullYear = -1;
        var lastMonth = -1;
        var lastRow = sheet.getLastRow();
        if (0 < lastRow - (headerInfo.idx_row + 1)) {
            var range = sheet.getRange(headerInfo.idx_row + 2, 1, lastRow - (headerInfo.idx_row + 1), Object.keys(CELL_HEADER_TITLES_DATA).length);
            if (range) {
                var valsRange = range.getValues();
                var rowNumBackuped = 0;
                var r = 0;
                for (; r < valsRange.length; r++) {
                    var row = valsRange[r];
                    var strCreatedAt = String(row[headerInfo.idx_createdAt]);
                    strCreatedAt = strCreatedAt.replace('(', '').replace(')', '').replace(' ', 'T') + "+09:00";
                    var dateCreatedAt = new Date(strCreatedAt);
                    if (g_datetime.getTime() - dateCreatedAt.getTime() <= DURATION_MILLISEC_NOT_FOR_BACKUP) {
                        break;
                    }
                    var year = dateCreatedAt.getFullYear();
                    var month = 1 + dateCreatedAt.getMonth();
                    if (lastMonth == -1) {
                        lastFullYear = year;
                        lastMonth = month;
                    }
                    else if (month != lastMonth) {
                        if (!moveData(nameBook, lastFullYear, lastMonth, sheetName, indexSheet, valsRange, rowNumBackuped, r - rowNumBackuped)) {
                            errOut("BACKUP PROCESS WAS TERMINATED.");
                            return;
                        }
                        lastFullYear = year;
                        lastMonth = month;
                        rowNumBackuped = r;
                    }
                }
                if (r - rowNumBackuped > 0) {
                    if (!moveData(nameBook, lastFullYear, lastMonth, sheetName, indexSheet, valsRange, rowNumBackuped, r - rowNumBackuped)) {
                        errOut("BACKUP PROCESS WAS TERMINATED.");
                        return;
                    }
                    rowNumBackuped = r;
                }
                if (rowNumBackuped > 0) {
                    sheet.deleteRows(headerInfo.idx_row + 2, rowNumBackuped);
                }
            }
        }
        indexSheet++;
    });
}
