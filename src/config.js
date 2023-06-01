/*
 * Version 1.00 made by yippym - 2023-05-02 19:00
 * https://github.com/Yippy/warp-tally-star-rail-sheet
 */
// Warp Tally Const
var WARP_TALLY_SHEET_SOURCE_REDIRECT_ID = '1hGTMKN1tJkTJOdxTtaEBSEgf1mg4JCYV-h1j9jufnuk';
var WARP_TALLY_SHEET_SUPPORTED_LOCALE = "en_GB";
var WARP_TALLY_SHEET_SCRIPT_VERSION = 1.00;
var WARP_TALLY_SHEET_SCRIPT_IS_ADD_ON = false;

// Auto Import Const
/* Add URL here to avoid showing on Sheet */
var AUTO_IMPORT_URL_FOR_API_BYPASS = ""; // Optional
var AUTO_IMPORT_BANNER_SETTINGS_FOR_IMPORT = {
  "Character Event Warp History": { "range_status": "E44", "range_toggle": "E37", "gacha_type": 11 },
  "Stellar Warp History": { "range_status": "E45", "range_toggle": "E38", "gacha_type": 1 },
  "Light Cone Event Warp History": { "range_status": "E46", "range_toggle": "E39", "gacha_type": 12 },
  "Departure Warp History": { "range_status": "E47", "range_toggle": "E40", "gacha_type": 2 },
};

var AUTO_IMPORT_LANGUAGE_SETTINGS_FOR_IMPORT = {
  "English": { "code": "en", "full_code": "en-us", "4_star": " (4-Star)", "5_star": " (5-Star)", "gacha_type_11": "Character Event Warp", "gacha_type_12": "Light Cone Event Warp", "gacha_type_1": "Stellar Warp", "gacha_type_2": "Departure Warp" },
  "German": { "code": "de", "full_code": "de-de", "4_star": " (4 Sterne)", "5_star": " (5 Sterne)", "gacha_type_11": "Figuren-Aktionswarp", "gacha_type_12": "Lichtkegel-Aktionswarp", "gacha_type_1": "Stellarwarp", "gacha_type_2": "Startwarp" },
  "French": { "code": "fr", "full_code": "fr-fr", "4_star": " (4★)", "5_star": " (5★)", "gacha_type_11": "Saut hyperespace événement de personnage", "gacha_type_12": "Saut hyperespace événement de cônes de lumière", "gacha_type_1": "Saut stellaire", "gacha_type_2": "Saut hyperespace de départ" },
  "Spanish": { "code": "es", "full_code": "es-es", "4_star": " (4★)", "5_star": " (5★)", "gacha_type_11": "Salto de evento de personaje", "gacha_type_12": "Salto de evento de cono de luz", "gacha_type_1": "Salto estelar", "gacha_type_2": "Salto de partida" },
  "Chinese Traditional": { "code": "zh-tw", "full_code": "zh-tw", "4_star": " (四星)", "5_star": " (五星)", "gacha_type_11": "角色活動躍遷", "gacha_type_12": "光錐活動躍遷", "gacha_type_1": "群星躍遷", "gacha_type_2": "始發躍遷" },
  "Chinese Simplified": { "code": "zh-cn", "full_code": "zh-cn", "4_star": " (四星)", "5_star": " (五星)", "gacha_type_11": "角色活动跃迁", "gacha_type_12": "光锥活动跃迁", "gacha_type_1": "群星跃迁", "gacha_type_2": "始发跃迁" },
  "Indonesian": { "code": "id", "full_code": "id-id", "4_star": " (4★)", "5_star": " (5★)", "gacha_type_11": "Event Warp Karakter", "gacha_type_12": "Event Warp Light Cone", "gacha_type_1": "Warp Bintang-Bintang", "gacha_type_2": "Warp Keberangkatan" },
  "Japanese": { "code": "ja", "full_code": "ja-jp", "4_star": " (★4)", "5_star": " (★5)", "gacha_type_11": "イベント跳躍・キャラクター", "gacha_type_12": "イベント跳躍・光円錐", "gacha_type_1": "群星跳躍", "gacha_type_2": "始発跳躍" },
  "Vietnamese": { "code": "vi", "full_code": "vi-vn", "4_star": " (4 sao)", "5_star": " (5 sao)", "gacha_type_11": "Bước Nhảy Sự Kiện Nhân Vật", "gacha_type_12": "Bước Nhảy Sự Kiện Nón Ánh Sáng", "gacha_type_1": "Bước Nhảy Chòm Sao", "gacha_type_2": "Bước Nhảy Đầu Tiên" },
  "Korean": { "code": "ko", "full_code": "ko-kr", "4_star": " (★4)", "5_star": " (★5)", "gacha_type_11": "캐릭터 이벤트 워프", "gacha_type_12": "광추 이벤트 워프", "gacha_type_1": "뭇별의 워프", "gacha_type_2": "초행길 워프" },
  "Portuguese": { "code": "pt", "full_code": "pt-pt", "4_star": " (4★)", "5_star": " (5★)", "gacha_type_11": "Salto Hiperespacial de Evento de Personagem", "gacha_type_12": "Salto Hiperespacial de Evento de Cone de Luz", "gacha_type_1": "Salto Hiperespacial Estelar", "gacha_type_2": "Salto Hiperespacial de Novatos" },
  "Thai": { "code": "th", "full_code": "th-th", "4_star": " (4 ดาว)", "5_star": " (5 ดาว)", "gacha_type_11": "กิจกรรมวาร์ปตัวละคร", "gacha_type_12": "กิจกรรมวาร์ป Light Cone", "gacha_type_1": "วาร์ปสู่ดวงดาว", "gacha_type_2": "ก้าวแรกแห่งการวาร์ป" },
  "Russian": { "code": "ru", "full_code": "ru-ru", "4_star": " (4★)", "5_star": " (5★)", "gacha_type_11": "Прыжок события: Персонаж", "gacha_type_12": "Прыжок события: Световой конус", "gacha_type_1": "Звёздный Прыжок", "gacha_type_2": "Отправной Прыжок" },
};

var AUTO_IMPORT_ADDITIONAL_QUERY = [
  "authkey_ver=1",
  "sign_type=2",
  "auth_appid=webview_gacha",
  "device_type=pc"
];

var AUTO_IMPORT_URL = "https://api-os-takumi.mihoyo.com/common/gacha_record/api/getGachaLog";
var AUTO_IMPORT_URL_CHINA = "https://api-takumi.mihoyo.com/event/gacha_info/api/getGachaLog";


var AUTO_IMPORT_URL_ERROR_CODE_AUTH_TIMEOUT = -101;
var AUTO_IMPORT_URL_ERROR_CODE_AUTH_INVALID = -100;
var AUTO_IMPORT_URL_ERROR_CODE_LANGUAGE_CODE = -108;
var AUTO_IMPORT_URL_ERROR_CODE_AUTHKEY_DENIED = -109;
var AUTO_IMPORT_URL_ERROR_CODE_REQUEST_PARAMS = -104;
var AUTO_IMPORT_URL_ERROR_CODE_VISIT_TOO_FREQUENTLY = -110;
var AUTO_IMPORT_WAIT_FOR_NEXT_API_CALL_IN_SECONDS = 1;

// Warp Tally Const
var WARP_TALLY_REDIRECT_SOURCE_ABOUT_SHEET_NAME = "About";
var WARP_TALLY_REDIRECT_SOURCE_AUTO_IMPORT_SHEET_NAME = "Auto Import";
var WARP_TALLY_REDIRECT_SOURCE_MAINTENANCE_SHEET_NAME = "Maintenance";
var WARP_TALLY_REDIRECT_SOURCE_HOYOLAB_SHEET_NAME = "HoYoLAB";
var WARP_TALLY_REDIRECT_SOURCE_BACKUP_SHEET_NAME = "Backup";
var WARP_TALLY_CHARACTER_EVENT_WARP_SHEET_NAME = "Character Event Warp History";
var WARP_TALLY_LIGHT_CONE_EVENT_WARP_SHEET_NAME = "Light Cone Event Warp History";
var WARP_TALLY_STELLAR_WARP_SHEET_NAME = "Stellar Warp History";
var WARP_TALLY_DEPARTURE_WARP_SHEET_NAME = "Departure Warp History";
var WARP_TALLY_WARP_HISTORY_SHEET_NAME = "Warp History";
var WARP_TALLY_SETTINGS_SHEET_NAME = "Settings";
var WARP_TALLY_DASHBOARD_SHEET_NAME = "Dashboard";
var WARP_TALLY_CHANGELOG_SHEET_NAME = "Changelog";
var WARP_TALLY_PITY_CHECKER_SHEET_NAME = "Pity Checker";

// Optional sheets
var WARP_TALLY_EVENTS_SHEET_NAME = "Events";
var WARP_TALLY_CHARACTERS_SHEET_NAME = "Characters";
var WARP_TALLY_LIGHT_CONES_SHEET_NAME = "Light Cones";
var WARP_TALLY_RESULTS_SHEET_NAME = "Results";
// Must match optional sheets names
var SETTINGS_FOR_OPTIONAL_SHEET = {
  "Events": {"setting_option": "B14"},
  "Results": {"setting_option": "B15"},
  "Characters": {"setting_option": "B16"},
  "Light Cones": {"setting_option": "B13"},
}

var WARP_TALLY_README_SHEET_NAME = "README";
var WARP_TALLY_AVAILABLE_SHEET_NAME = "Available";
var WARP_TALLY_SHARD_CALCULATOR_SHEET_NAME = "Shard Calculator";
var WARP_TALLY_ALL_WARP_HISTORY_SHEET_NAME = "All Warp History";
var WARP_TALLY_ITEMS_SHEET_NAME = "Items";
var WARP_TALLY_NAME_OF_WARP_HISTORY = [WARP_TALLY_CHARACTER_EVENT_WARP_SHEET_NAME, WARP_TALLY_STELLAR_WARP_SHEET_NAME, WARP_TALLY_LIGHT_CONE_EVENT_WARP_SHEET_NAME, WARP_TALLY_DEPARTURE_WARP_SHEET_NAME];

// Import Const
var IMPORT_STATUS_COMPLETE = "COMPLETE";
var IMPORT_STATUS_FAILED = "FAILED";
var IMPORT_STATUS_WARP_HISTORY_COMPLETE = "DONE";
var IMPORT_STATUS_WARP_HISTORY_NOT_FOUND = "NOT FOUND";
var IMPORT_STATUS_WARP_HISTORY_EMPTY = "EMPTY";

// Scheduler Const
var SCHEDULER_TRIGGER_ON_TEXT = "ON";
var SCHEDULER_TRIGGER_OFF_TEXT = "OFF";
var SCHEDULER_RUN_AT_WHICH_DAY = {
  "Monday": ScriptApp.WeekDay.MONDAY,
  "Tuesday": ScriptApp.WeekDay.TUESDAY,
  "Wednesday": ScriptApp.WeekDay.WEDNESDAY,
  "Thursday": ScriptApp.WeekDay.THURSDAY,
  "Friday": ScriptApp.WeekDay.FRIDAY,
  "Saturday": ScriptApp.WeekDay.SATURDAY,
  "Sunday": ScriptApp.WeekDay.SUNDAY
};
var SCHEDULER_RUN_AT_HOUR = {
  "Run at 1:00": 1,
  "Run at 2:00": 2,
  "Run at 3:00": 3,
  "Run at 4:00": 4,
  "Run at 5:00": 5,
  "Run at 6:00": 6,
  "Run at 7:00": 7,
  "Run at 8:00": 8,
  "Run at 9:00": 9,
  "Run at 10:00": 10,
  "Run at 11:00": 11,
  "Run at 12:00": 12,
  "Run at 13:00": 13,
  "Run at 14:00": 14,
  "Run at 15:00": 15,
  "Run at 16:00": 16,
  "Run at 17:00": 17,
  "Run at 18:00": 18,
  "Run at 19:00": 19,
  "Run at 20:00": 20,
  "Run at 21:00": 21,
  "Run at 22:00": 22,
  "Run at 23:00": 23,
  "Run at Midnight": 0
};
var SCHEDULER_RUN_AT_EVERY_HOUR = {
  "Every hour": 1,
  "Every 2 hours": 2,
  "Every 3 hours": 3,
  "Every 4 hours": 4,
  "Every 5 hours": 5,
  "Every 6 hours": 6,
  "Every 7 hours": 7,
  "Every 8 hours": 8,
  "Every 9 hours": 9,
  "Every 10 hours": 10,
  "Every 11 hours": 11,
  "Every 12 hours": 12,
  "Every 13 hours": 13,
  "Every 14 hours": 14,
  "Every 15 hours": 15,
  "Every 16 hours": 16,
  "Every 17 hours": 17,
  "Every 18 hours": 18,
  "Every 19 hours": 19,
  "Every 20 hours": 20,
  "Every 21 hours": 21,
  "Every 22 hours": 22,
  "Every 23 hours": 23,
  "Every 24 hours": 24
};

const CACHED_AUTHKEY_TIMEOUT = 1000 * 60 * 60 * 24;
