/**
 * ISLAND: LAST SURVIVOR - Google Apps Script Backend
 * オンライン協力型無人島サバイバルRPG
 */

// ==========================================
// データベース設定 (スプレッドシートID)
// ==========================================
// ※ スプレッドシートIDを設定してください。
// 空白文字 '' の場合は、コンテナバインドスクリプトのアクティブスプレッドシートを自動参照します。
var SPREADSHEET_ID = '';

/**
 * スプレッドシートオブジェクトを取得するヘルパー関数
 */
function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== '') {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  // スタンドアロンまたはプロパティに設定されている場合
  var propId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (propId) {
    return SpreadsheetApp.openById(propId);
  }
  // コンテナバインドスクリプトの場合
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Web App エントリポイント
 * index.html（CSS・JS完全内包版）を直接返すため、テンプレート評価の依存性や読み込み失敗を完全排除
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ISLAND: LAST SURVIVOR')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * テンプレートインクルード用ヘルパー（後方互換性のため維持）
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (err) {
    return '<!-- include failed: ' + filename + ' -->';
  }
}

// ==========================================
// シート定義と初期化
// ==========================================
var SHEETS = {
  PLAYERS: 'Players',
  ISLANDS: 'Islands',
  MEMBERS: 'Members',
  INVENTORY: 'Inventory',
  SHARED_INVENTORY: 'SharedInventory',
  ITEMS: 'Items',
  RECIPES: 'Recipes',
  EVENTS: 'Events',
  LOGS: 'Logs',
  MESSAGES: 'Messages',
  BUILDINGS: 'Buildings',
  ACHIEVEMENTS: 'Achievements'
};

var HEADERS = {
  Players: ['player_id', 'name', 'role', 'hp', 'hunger', 'thirst', 'stamina', 'status', 'created_at', 'updated_at', 'auth_token'],
  Islands: ['island_id', 'island_name', 'invite_code', 'host_player_id', 'day', 'time_period', 'island_status', 'created_at', 'updated_at'],
  Members: ['island_id', 'player_id', 'role', 'member_status', 'joined_at', 'last_seen'],
  Inventory: ['player_id', 'item_id', 'quantity'],
  SharedInventory: ['island_id', 'item_id', 'quantity'],
  Items: ['item_id', 'item_name', 'category', 'description', 'effect'],
  Recipes: ['recipe_id', 'result_item_id', 'wood', 'stone', 'other_material', 'description'],
  Events: ['event_id', 'event_name', 'description', 'effect_type', 'effect_value'],
  Logs: ['log_id', 'island_id', 'player_id', 'day', 'action', 'created_at'],
  Messages: ['message_id', 'island_id', 'player_id', 'player_name', 'message', 'created_at'],
  Buildings: ['building_id', 'island_id', 'building_type', 'level', 'created_by', 'created_at'],
  Achievements: ['achievement_id', 'player_id', 'achievement_name', 'unlocked_at']
};

/**
 * データベースの初期設定（12シート作成、ヘッダー設定、初期マスターデータ投入）
 * 既存のデータを破壊せず、何度実行しても安全に重複追加を防ぐ設計。
 */
function setupDatabase() {
  var ss = getSpreadsheet();
  
  // 1. シートとヘッダーの自動生成
  for (var key in HEADERS) {
    var sheetName = key;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    // ヘッダーが空なら設定
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS[key]);
      sheet.getRange(1, 1, 1, HEADERS[key].length).setFontWeight('bold').setBackground('#2d6a4f').setFontColor('#ffffff');
    }
  }

  // デフォルトのシート1が残っていれば整理（任意）
  var defaultSheet = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch(e) {}
  }

  // 2. アイテムマスターデータの投入
  var itemsSheet = ss.getSheetByName(SHEETS.ITEMS);
  var existingItems = itemsSheet.getDataRange().getValues();
  var itemMap = {};
  for (var i = 1; i < existingItems.length; i++) {
    itemMap[existingItems[i][0]] = true;
  }

  var defaultItems = [
    ['wood', '木材', 'material', '建築やクラフトに使用する基本的な木', 'なし'],
    ['stone', '石', 'material', '道具や建物の土台に使う頑丈な石', 'なし'],
    ['berry', '木の実', 'food', '空腹を満たす野いちご', 'hunger:20,thirst:5'],
    ['water', '真水', 'drink', '生存に不可欠なきれいな淡水', 'thirst:35'],
    ['fish', '魚', 'food', '海岸で捕れた魚。栄養満点', 'hunger:35,hp:10'],
    ['herb', '薬草', 'medical', '傷を癒やす貴重な野草', 'hp:25'],
    ['iron_ore', '鉄鉱石', 'material', '洞窟深くで見つかる頑丈な鉱石', 'なし'],
    ['rope', 'ロープ', 'material', '草や皮を編み込んだ丈夫な綱', 'なし'],
    ['cloth', '布', 'material', 'テントや帆を作るための布', 'なし'],
    ['stone_axe', '石の斧', 'tool', '木材の収集効率を向上させる', 'wood_boost:1'],
    ['spear', '槍', 'tool', '野生動物の撃退や狩りに有利', 'hunt_boost:1'],
    ['fishing_rod', '釣り竿', 'tool', '海岸での魚の獲得確率UP', 'fish_boost:1'],
    ['torch', 'たいまつ', 'tool', '洞窟探索時の事故を防止し夜を照らす', 'light:1'],
    ['raft_piece', 'いかだの骨組み', 'special', '脱出用いかだの重要部品', 'escape_material:1']
  ];

  for (var j = 0; j < defaultItems.length; j++) {
    if (!itemMap[defaultItems[j][0]]) {
      itemsSheet.appendRow(defaultItems[j]);
    }
  }

  // 3. レシピマスターデータの投入
  var recipesSheet = ss.getSheetByName(SHEETS.RECIPES);
  var existingRecipes = recipesSheet.getDataRange().getValues();
  var recipeMap = {};
  for (var r = 1; r < existingRecipes.length; r++) {
    recipeMap[existingRecipes[r][0]] = true;
  }

  var defaultRecipes = [
    ['rec_axe', 'stone_axe', 3, 2, '', '木材×3, 石×2 で作成する石斧'],
    ['rec_spear', 'spear', 3, 1, 'rope:1', '木材×3, 石×1, ロープ×1 で作る狩猟用槍'],
    ['rec_rod', 'fishing_rod', 2, 0, 'rope:2', '木材×2, ロープ×2 で作る釣り竿'],
    ['rec_torch', 'torch', 1, 0, 'cloth:1', '木材×1, 布×1 で作るたいまつ'],
    ['rec_rope', 'rope', 0, 0, 'herb:2', '薬草/繊維×2 で編み上げるロープ'],
    ['rec_cloth', 'cloth', 0, 0, 'herb:3', '繊維×3 から織り上げる布'],
    ['rec_raft', 'raft_piece', 10, 5, 'rope:5,cloth:3', '木材×10, 石×5, ロープ×5, 布×3 で作るいかだ部品']
  ];

  for (var k = 0; k < defaultRecipes.length; k++) {
    if (!recipeMap[defaultRecipes[k][0]]) {
      recipesSheet.appendRow(defaultRecipes[k]);
    }
  }

  // 4. イベントマスターデータの投入
  var eventsSheet = ss.getSheetByName(SHEETS.EVENTS);
  var existingEvents = eventsSheet.getDataRange().getValues();
  var eventMap = {};
  for (var e = 1; e < existingEvents.length; e++) {
    eventMap[existingEvents[e][0]] = true;
  }

  var defaultEvents = [
    ['STORM', '暴風雨の襲来', '激しい嵐が島を襲い、体力を奪う。', 'damage_all', '15'],
    ['WILD_BEAST', '凶暴な野獣の接近', '唸り声を上げる獣が迫る。', 'encounter_beast', '20'],
    ['FIND_FOOD', '豊かな実りの発見', '野生の果樹園を発見した！', 'reward_food', 'berry:4,water:2'],
    ['TREASURE_CHEST', '漂着した難破船の宝箱', '古いトランクが流れ着いていた！', 'reward_rare', 'iron_ore:2,cloth:2,rope:2'],
    ['CAVE_ACCIDENT', '洞窟の落石事故', '足元が崩れ落石に巻き込まれた。', 'damage_self', '25'],
    ['RARE_RESOURCE', '鉱脈の発見', '洞窟奥深くで高純度の鉄鉱石鉱脈を発見した。', 'reward_ore', 'iron_ore:4'],
    ['RESCUE_CRATE', '救護物資の漂着', '医療マークの入った木箱を発見した。', 'reward_med', 'herb:4,water:3']
  ];

  for (var m = 0; m < defaultEvents.length; m++) {
    if (!eventMap[defaultEvents[m][0]]) {
      eventsSheet.appendRow(defaultEvents[m]);
    }
  }

  return { success: true, message: 'データベースが正常に初期化されました。' };
}

// ==========================================
// ユーティリティ & 認証関数
// ==========================================

function generateUUID() {
  return Utilities.getUuid();
}

function generateInviteCode() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // 誤読しやすいO,0,I,1を除く
  var code = '';
  for (var i = 0; i < 6; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}

function getNowString() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}

/**
 * プレイヤー認証（player_id と auth_token の厳密な検証）
 */
function authenticatePlayer(ss, playerId, token) {
  if (!playerId || !token) {
    throw new Error('認証情報が不足しています。');
  }
  var sheet = ss.getSheetByName(SHEETS.PLAYERS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === playerId) {
      var dbToken = data[i][10];
      if (dbToken === token) {
        return {
          row: i + 1,
          playerId: data[i][0],
          player_id: data[i][0],
          name: data[i][1],
          role: data[i][2],
          hp: Number(data[i][3]),
          hunger: Number(data[i][4]),
          thirst: Number(data[i][5]),
          stamina: Number(data[i][6]),
          status: data[i][7]
        };
      } else {
        throw new Error('セッションが無効です。再ログインしてください。');
      }
    }
  }
  throw new Error('プレイヤーが見つかりません。');
}

/**
 * プレイヤーが指定の島に所属しているか検証
 */
function verifyIslandMembership(ss, islandId, playerId) {
  var membersSheet = ss.getSheetByName(SHEETS.MEMBERS);
  var data = membersSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === islandId && data[i][1] === playerId) {
      if (data[i][3] === 'LEFT') {
        throw new Error('この島から退出済みです。再参加してください。');
      }
      return { row: i + 1, member_status: data[i][3], role: data[i][2] };
    }
  }
  throw new Error('この島への参加権限がありません。');
}

/**
 * 島の詳細を取得
 */
function getIslandData(ss, islandId) {
  var sheet = ss.getSheetByName(SHEETS.ISLANDS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === islandId) {
      return {
        row: i + 1,
        island_id: data[i][0],
        island_name: data[i][1],
        invite_code: data[i][2],
        host_player_id: data[i][3],
        day: Number(data[i][4]),
        time_period: data[i][5],
        island_status: data[i][6]
      };
    }
  }
  throw new Error('指定された島が見つかりません。');
}

/**
 * ログを追加
 */
function addLog(ss, islandId, playerId, day, action) {
  var sheet = ss.getSheetByName(SHEETS.LOGS);
  sheet.appendRow([generateUUID(), islandId, playerId, day, action, getNowString()]);
}

// ==========================================
// STEP 1: プレイヤー登録 / セッション検証 / セーブ＆ロード
// ==========================================

function registerPlayer(name, role) {
  if (!name || name.trim() === '') {
    return { success: false, message: '名前を入力してください。' };
  }
  var validRoles = ['EXPLORER', 'HUNTER', 'CRAFTSMAN', 'MEDIC'];
  if (validRoles.indexOf(role) === -1) {
    role = 'EXPLORER';
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var playersSheet = ss.getSheetByName(SHEETS.PLAYERS);
    var playerId = 'p_' + generateUUID().substring(0, 8);
    var token = 'tok_' + generateUUID();
    var now = getNowString();

    // 初期ステータス
    var hp = 100;
    var hunger = 100;
    var thirst = 100;
    var stamina = 100;
    var status = 'ALIVE';

    playersSheet.appendRow([playerId, name.trim(), role, hp, hunger, thirst, stamina, status, now, now, token]);

    // 初期インベントリの付与
    var invSheet = ss.getSheetByName(SHEETS.INVENTORY);
    invSheet.appendRow([playerId, 'stone_axe', 1]);
    invSheet.appendRow([playerId, 'water', 2]);
    invSheet.appendRow([playerId, 'berry', 3]);

    return {
      success: true,
      playerId: playerId,
      token: token,
      player: {
        playerId: playerId,
        name: name.trim(),
        role: role,
        hp: hp,
        hunger: hunger,
        thirst: thirst,
        stamina: stamina,
        status: status
      }
    };
  } catch (err) {
    return { success: false, message: '登録に失敗しました: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * セッショントークン検証と自動再ログイン
 */
function verifySession(playerId, token) {
  try {
    var ss = getSpreadsheet();
    var player = authenticatePlayer(ss, playerId, token);
    
    // 現在所属している島を検索
    var membersSheet = ss.getSheetByName(SHEETS.MEMBERS);
    var mData = membersSheet.getDataRange().getValues();
    var currentIslandId = null;
    for (var i = 1; i < mData.length; i++) {
      if (mData[i][1] === playerId && mData[i][3] === 'ACTIVE') {
        currentIslandId = mData[i][0];
        break;
      }
    }

    var island = null;
    if (currentIslandId) {
      try {
        island = getIslandData(ss, currentIslandId);
      } catch (e) {
        currentIslandId = null;
      }
    }

    return {
      success: true,
      player: player,
      currentIsland: island
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ==========================================
// STEP 2: 島の作成、招待コード、フレンド参加、退出
// ==========================================

function createIsland(playerId, token, islandName) {
  if (!islandName || islandName.trim() === '') {
    islandName = '名もなき孤島';
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var player = authenticatePlayer(ss, playerId, token);

    var islandId = 'isl_' + generateUUID().substring(0, 8);
    var inviteCode = generateInviteCode();
    var now = getNowString();

    var islandsSheet = ss.getSheetByName(SHEETS.ISLANDS);
    islandsSheet.appendRow([
      islandId,
      islandName.trim(),
      inviteCode,
      playerId,
      1,           // day 1
      'MORNING',   // 朝
      'ACTIVE',    // ステータス
      now,
      now
    ]);

    // ホストをメンバーに追加
    var membersSheet = ss.getSheetByName(SHEETS.MEMBERS);
    membersSheet.appendRow([islandId, playerId, player.role, 'ACTIVE', now, now]);

    // 共有倉庫の初期化（初期物資）
    var sharedSheet = ss.getSheetByName(SHEETS.SHARED_INVENTORY);
    sharedSheet.appendRow([islandId, 'wood', 10]);
    sharedSheet.appendRow([islandId, 'stone', 5]);
    sharedSheet.appendRow([islandId, 'water', 5]);
    sharedSheet.appendRow([islandId, 'berry', 5]);

    // ログ記録
    addLog(ss, islandId, playerId, 1, player.name + ' が新しい島【' + islandName.trim() + '】を開拓しました！');

    return {
      success: true,
      islandId: islandId,
      inviteCode: inviteCode,
      islandName: islandName.trim()
    };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function joinIsland(playerId, token, inviteCode) {
  if (!inviteCode || inviteCode.trim() === '') {
    return { success: false, message: '招待コードを入力してください。' };
  }
  inviteCode = inviteCode.trim().toUpperCase();

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var player = authenticatePlayer(ss, playerId, token);

    // 招待コードの照合
    var islandsSheet = ss.getSheetByName(SHEETS.ISLANDS);
    var iData = islandsSheet.getDataRange().getValues();
    var targetIsland = null;
    for (var i = 1; i < iData.length; i++) {
      if (String(iData[i][2]).toUpperCase() === inviteCode) {
        if (iData[i][6] !== 'ACTIVE') {
          return { success: false, message: 'この島は既にゲームが終了しています。' };
        }
        targetIsland = {
          row: i + 1,
          island_id: iData[i][0],
          island_name: iData[i][1],
          day: iData[i][4]
        };
        break;
      }
    }

    if (!targetIsland) {
      return { success: false, message: '有効な招待コードが見つかりません。' };
    }

    var islandId = targetIsland.island_id;
    var membersSheet = ss.getSheetByName(SHEETS.MEMBERS);
    var mData = membersSheet.getDataRange().getValues();
    
    var activeMemberCount = 0;
    var existingMemberRow = -1;

    for (var m = 1; m < mData.length; m++) {
      if (mData[m][0] === islandId) {
        if (mData[m][3] === 'ACTIVE') {
          activeMemberCount++;
        }
        if (mData[m][1] === playerId) {
          existingMemberRow = m + 1;
        }
      }
    }

    var now = getNowString();

    if (existingMemberRow > 0) {
      // 既存メンバーの再参加
      membersSheet.getRange(existingMemberRow, 4).setValue('ACTIVE');
      membersSheet.getRange(existingMemberRow, 6).setValue(now);
    } else {
      // 新規参加の場合、最大4人上限チェック
      if (activeMemberCount >= 4) {
        return { success: false, message: '島は満員です（最大4人まで参加可能）。' };
      }
      membersSheet.appendRow([islandId, playerId, player.role, 'ACTIVE', now, now]);
    }

    addLog(ss, islandId, playerId, targetIsland.day, player.name + ' が島に合流しました。');

    return {
      success: true,
      islandId: islandId,
      islandName: targetIsland.island_name
    };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function leaveIsland(playerId, token, islandId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var player = authenticatePlayer(ss, playerId, token);
    var membership = verifyIslandMembership(ss, islandId, playerId);

    var membersSheet = ss.getSheetByName(SHEETS.MEMBERS);
    membersSheet.getRange(membership.row, 4).setValue('LEFT');

    var island = getIslandData(ss, islandId);
    addLog(ss, islandId, playerId, island.day, player.name + ' が島から一時退出しました。');

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// STEP 3: 島のメンバーと状態の同期 (3〜5秒ポーリング用)
// ==========================================

function getIslandState(playerId, token, islandId, lastMessageTime, lastLogTime) {
  try {
    var ss = getSpreadsheet();
    var me = authenticatePlayer(ss, playerId, token);
    var membership = verifyIslandMembership(ss, islandId, playerId);

    var now = new Date();
    var nowStr = getNowString();

    // 最終通信日時更新 (last_seen)
    var membersSheet = ss.getSheetByName(SHEETS.MEMBERS);
    membersSheet.getRange(membership.row, 6).setValue(nowStr);

    // 1. 島情報
    var island = getIslandData(ss, islandId);

    // 2. メンバー一覧とプレイヤーデータ
    var mData = membersSheet.getDataRange().getValues();
    var islandMemberIds = [];
    var memberRoles = {};
    for (var m = 1; m < mData.length; m++) {
      if (mData[m][0] === islandId && mData[m][3] === 'ACTIVE') {
        islandMemberIds.push(mData[m][1]);
        memberRoles[mData[m][1]] = {
          role: mData[m][2],
          last_seen: mData[m][5]
        };
      }
    }

    var playersSheet = ss.getSheetByName(SHEETS.PLAYERS);
    var pData = playersSheet.getDataRange().getValues();
    var members = [];
    for (var p = 1; p < pData.length; p++) {
      var pId = pData[p][0];
      if (islandMemberIds.indexOf(pId) !== -1) {
        var lastSeenDate = new Date(memberRoles[pId].last_seen);
        var isOnline = (now - lastSeenDate) < 20000; // 20秒以内ならオンライン
        
        var memberObj = {
          playerId: pId,
          name: pData[p][1],
          role: pData[p][2],
          hp: Number(pData[p][3]),
          hunger: Number(pData[p][4]),
          thirst: Number(pData[p][5]),
          stamina: Number(pData[p][6]),
          status: pData[p][7],
          isOnline: isOnline
        };
        members.push(memberObj);

        // 自身の最新ステータスを更新
        if (pId === playerId) {
          me = memberObj;
        }
      }
    }

    // 3. 個人インベントリ
    var invSheet = ss.getSheetByName(SHEETS.INVENTORY);
    var invData = invSheet.getDataRange().getValues();
    var personalInventory = [];
    for (var iv = 1; iv < invData.length; iv++) {
      if (invData[iv][0] === playerId) {
        personalInventory.push({
          itemId: invData[iv][1],
          quantity: Number(invData[iv][2])
        });
      }
    }

    // 4. 共有倉庫
    var sharedSheet = ss.getSheetByName(SHEETS.SHARED_INVENTORY);
    var sData = sharedSheet.getDataRange().getValues();
    var sharedInventory = [];
    for (var s = 1; s < sData.length; s++) {
      if (sData[s][0] === islandId) {
        sharedInventory.push({
          itemId: sData[s][1],
          quantity: Number(sData[s][2])
        });
      }
    }

    // 5. 建築物
    var buildSheet = ss.getSheetByName(SHEETS.BUILDINGS);
    var bData = buildSheet.getDataRange().getValues();
    var buildings = [];
    for (var b = 1; b < bData.length; b++) {
      if (bData[b][1] === islandId) {
        buildings.push({
          buildingId: bData[b][0],
          buildingType: bData[b][2],
          level: Number(bData[b][3]),
          createdBy: bData[b][4],
          createdAt: bData[b][5]
        });
      }
    }

    // 6. チャット新着取得 (最新25件)
    var msgSheet = ss.getSheetByName(SHEETS.MESSAGES);
    var msgData = msgSheet.getDataRange().getValues();
    var messages = [];
    var startMsg = Math.max(1, msgData.length - 25);
    for (var mg = startMsg; mg < msgData.length; mg++) {
      if (msgData[mg][1] === islandId) {
        messages.push({
          messageId: msgData[mg][0],
          playerId: msgData[mg][2],
          playerName: msgData[mg][3],
          message: msgData[mg][4],
          createdAt: msgData[mg][5]
        });
      }
    }

    // 7. 活動ログ新着取得 (最新25件)
    var logSheet = ss.getSheetByName(SHEETS.LOGS);
    var logData = logSheet.getDataRange().getValues();
    var logs = [];
    var startLog = Math.max(1, logData.length - 25);
    for (var lg = startLog; lg < logData.length; lg++) {
      if (logData[lg][1] === islandId) {
        logs.push({
          logId: logData[lg][0],
          playerId: logData[lg][2],
          day: logData[lg][3],
          action: logData[lg][4],
          createdAt: logData[lg][5]
        });
      }
    }

    // アイテム定義マスター（クライアント表示用）
    var itemsSheet = ss.getSheetByName(SHEETS.ITEMS);
    var itemsRaw = itemsSheet.getDataRange().getValues();
    var itemsMap = {};
    for (var it = 1; it < itemsRaw.length; it++) {
      itemsMap[itemsRaw[it][0]] = {
        name: itemsRaw[it][1],
        category: itemsRaw[it][2],
        description: itemsRaw[it][3],
        effect: itemsRaw[it][4]
      };
    }

    return {
      success: true,
      island: island,
      player: me,
      members: members,
      personalInventory: personalInventory,
      sharedInventory: sharedInventory,
      buildings: buildings,
      messages: messages,
      logs: logs,
      itemsMap: itemsMap
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ==========================================
// STEP 4: 探索、時間進行、サバイバル、共有倉庫
// ==========================================

/**
 * 探索システム
 */
function exploreLocation(playerId, token, islandId, location) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var player = authenticatePlayer(ss, playerId, token);
    verifyIslandMembership(ss, islandId, playerId);
    var island = getIslandData(ss, islandId);

    if (island.island_status !== 'ACTIVE') {
      return { success: false, message: 'ゲームは既に終了しています。' };
    }
    if (player.status === 'DEAD' || player.hp <= 0) {
      return { success: false, message: '死亡状態では探索できません。仲間に救助を求めてください。' };
    }

    // 役割によるスタミナ消費計算
    // EXPLORER は消費30%軽減
    var staminaCost = 20;
    if (player.role === 'EXPLORER') {
      staminaCost = 14;
    }

    if (player.stamina < staminaCost) {
      return { success: false, message: 'スタミナが不足しています（必要: ' + staminaCost + '）。休憩または食事をしてください。' };
    }

    var newStamina = Math.max(0, player.stamina - staminaCost);
    var playersSheet = ss.getSheetByName(SHEETS.PLAYERS);
    playersSheet.getRange(player.row, 7).setValue(newStamina);

    // 探索ドロップ抽選
    var foundItems = [];
    var logMsg = '';
    var eventTriggered = null;

    var roll = Math.random();
    var isExplorer = (player.role === 'EXPLORER');
    var isHunter = (player.role === 'HUNTER');

    if (location === 'forest') {
      // 森: 木材, 木の実, 薬草, 野生動物遭遇
      var woodQty = isExplorer ? 3 + Math.floor(Math.random() * 3) : 2 + Math.floor(Math.random() * 2);
      foundItems.push({ itemId: 'wood', quantity: woodQty });

      var berryQty = isHunter ? 4 : 2;
      foundItems.push({ itemId: 'berry', quantity: berryQty });

      if (Math.random() < (isExplorer ? 0.6 : 0.4)) {
        foundItems.push({ itemId: 'herb', quantity: 1 + Math.floor(Math.random() * 2) });
      }

      if (Math.random() < 0.25) {
        // 野生動物遭遇イベント
        if (isHunter) {
          foundItems.push({ itemId: 'fish', quantity: 2 }); // 肉代わり
          logMsg = '森で野生の獣と遭遇したが、ハンターの技で仕留めた！';
        } else {
          var beastDmg = 15;
          var newHp = Math.max(0, player.hp - beastDmg);
          playersSheet.getRange(player.row, 4).setValue(newHp);
          if (newHp === 0) {
            playersSheet.getRange(player.row, 8).setValue('DEAD');
          }
          logMsg = '森で野獣に襲われ、HPを ' + beastDmg + ' 失った！';
        }
      }
    } else if (location === 'beach') {
      // 海岸: 流木, 貝・魚, 水
      var driftWood = 1 + Math.floor(Math.random() * 2);
      foundItems.push({ itemId: 'wood', quantity: driftWood });

      var fishQty = isHunter ? 3 : (Math.random() < 0.6 ? 2 : 1);
      foundItems.push({ itemId: 'fish', quantity: fishQty });

      if (Math.random() < 0.5) {
        foundItems.push({ itemId: 'stone', quantity: 2 });
      }
      if (Math.random() < 0.4) {
        foundItems.push({ itemId: 'water', quantity: 1 });
      }

      // レア宝箱イベント
      if (Math.random() < (isExplorer ? 0.35 : 0.15)) {
        foundItems.push({ itemId: 'cloth', quantity: 1 });
        foundItems.push({ itemId: 'rope', quantity: 1 });
        logMsg = '海岸に流れ着いた救援コンテナを発見！布とロープを入手した！';
      }
    } else if (location === 'cave') {
      // 洞窟: 石, 鉱石, レア素材, 落石事故
      var stoneQty = 3 + Math.floor(Math.random() * 3);
      foundItems.push({ itemId: 'stone', quantity: stoneQty });

      if (Math.random() < (isExplorer ? 0.7 : 0.4)) {
        foundItems.push({ itemId: 'iron_ore', quantity: isExplorer ? 2 : 1 });
      }

      // 洞窟事故判定
      if (Math.random() < (isExplorer ? 0.15 : 0.3)) {
        var caveDmg = 20;
        var newHpCave = Math.max(0, player.hp - caveDmg);
        playersSheet.getRange(player.row, 4).setValue(newHpCave);
        if (newHpCave === 0) {
          playersSheet.getRange(player.row, 8).setValue('DEAD');
        }
        logMsg = '暗闇の洞窟で落石事故に巻き込まれ、HPを ' + caveDmg + ' 消耗した！';
      }
    }

    // 獲得アイテムを個人インベントリに追加
    var invSheet = ss.getSheetByName(SHEETS.INVENTORY);
    var invData = invSheet.getDataRange().getValues();

    for (var f = 0; f < foundItems.length; f++) {
      var itemObj = foundItems[f];
      var foundRow = -1;
      for (var r = 1; r < invData.length; r++) {
        if (invData[r][0] === playerId && invData[r][1] === itemObj.itemId) {
          foundRow = r + 1;
          break;
        }
      }
      if (foundRow > 0) {
        var currentQty = Number(invSheet.getRange(foundRow, 3).getValue());
        invSheet.getRange(foundRow, 3).setValue(currentQty + itemObj.quantity);
      } else {
        invSheet.appendRow([playerId, itemObj.itemId, itemObj.quantity]);
      }
    }

    // ログ記録
    var itemsStr = foundItems.map(function(fi) { return fi.itemId + '×' + fi.quantity; }).join(', ');
    var locationNames = { forest: '森', beach: '海岸', cave: '洞窟' };
    var baseLog = player.name + ' が【' + locationNames[location] + '】を探索し、[' + itemsStr + '] を入手しました。';
    if (logMsg) {
      baseLog += ' (' + logMsg + ')';
    }
    addLog(ss, islandId, playerId, island.day, baseLog);

    player.stamina = newStamina;
    SpreadsheetApp.flush();

    return {
      success: true,
      foundItems: foundItems,
      player: player,
      message: baseLog
    };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 時間進行システム（朝・昼・夕・夜・日数）
 * 排他制御で二重進行を厳密に防止
 */
function advanceTime(playerId, token, islandId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var player = authenticatePlayer(ss, playerId, token);
    verifyIslandMembership(ss, islandId, playerId);
    var island = getIslandData(ss, islandId);

    if (island.island_status !== 'ACTIVE') {
      return { success: false, message: 'ゲームは終了しています。' };
    }

    var timePeriods = ['MORNING', 'NOON', 'EVENING', 'NIGHT'];
    var currentIdx = timePeriods.indexOf(island.time_period);
    var nextIdx = (currentIdx + 1) % 4;
    var nextTime = timePeriods[nextIdx];
    var nextDay = island.day;
    var isNewDay = (nextIdx === 0);

    if (isNewDay) {
      nextDay++;
    }

    // 島データの更新
    var islandsSheet = ss.getSheetByName(SHEETS.ISLANDS);
    islandsSheet.getRange(island.row, 5).setValue(nextDay);
    islandsSheet.getRange(island.row, 6).setValue(nextTime);
    islandsSheet.getRange(island.row, 9).setValue(getNowString());

    // 建築物による効果確認
    var buildSheet = ss.getSheetByName(SHEETS.BUILDINGS);
    var bData = buildSheet.getDataRange().getValues();
    var hasCampfire = false;
    var hasTent = false;
    var hasCabin = false;
    var hasWaterCollector = false;

    for (var b = 1; b < bData.length; b++) {
      if (bData[b][1] === islandId) {
        var bType = bData[b][2];
        if (bType === 'campfire') hasCampfire = true;
        if (bType === 'tent') hasTent = true;
        if (bType === 'cabin') hasCabin = true;
        if (bType === 'water_collector') hasWaterCollector = true;
      }
    }

    // 貯水設備があれば翌朝に水を共有倉庫に追加
    if (isNewDay && hasWaterCollector) {
      addToSharedInventory(ss, islandId, 'water', 4);
      addLog(ss, islandId, 'SYSTEM', nextDay, '貯水設備により、共有倉庫にきれいな水×4が蓄えられました。');
    }

    // メンバー全員のステータス進行（空腹・水分減少・スタミナ回復等）
    var membersSheet = ss.getSheetByName(SHEETS.MEMBERS);
    var mData = membersSheet.getDataRange().getValues();
    var playersSheet = ss.getSheetByName(SHEETS.PLAYERS);
    var pData = playersSheet.getDataRange().getValues();

    var allDead = true;

    for (var m = 1; m < mData.length; m++) {
      if (mData[m][0] === islandId && mData[m][3] === 'ACTIVE') {
        var memPlayerId = mData[m][1];
        for (var p = 1; p < pData.length; p++) {
          if (pData[p][0] === memPlayerId) {
            var curHp = Number(pData[p][3]);
            var curHunger = Number(pData[p][4]);
            var curThirst = Number(pData[p][5]);
            var curStamina = Number(pData[p][6]);
            var curStatus = pData[p][7];

            if (curStatus === 'ALIVE') {
              // 空腹度 -10, 水分 -15
              var newHunger = Math.max(0, curHunger - 10);
              var newThirst = Math.max(0, curThirst - 15);

              // スタミナ回復（拠点設備による向上）
              var stamRecover = 15;
              if (hasCabin) stamRecover += 25;
              else if (hasTent) stamRecover += 15;
              var newStamina = Math.min(100, curStamina + stamRecover);

              // 飢餓・脱水ダメージ
              var damage = 0;
              if (newHunger === 0) damage += 15;
              if (newThirst === 0) damage += 20;

              // 夜間の寒さ（焚き火なしだとダメージ）
              if (island.time_period === 'NIGHT' && !hasCampfire) {
                damage += 10;
              }

              var newHp = Math.max(0, curHp - damage);
              var newStatus = (newHp === 0) ? 'DEAD' : 'ALIVE';

              if (newStatus === 'ALIVE') {
                allDead = false;
              }

              playersSheet.getRange(p + 1, 4).setValue(newHp);
              playersSheet.getRange(p + 1, 5).setValue(newHunger);
              playersSheet.getRange(p + 1, 6).setValue(newThirst);
              playersSheet.getRange(p + 1, 7).setValue(newStamina);
              playersSheet.getRange(p + 1, 8).setValue(newStatus);
            }
          }
        }
      }
    }

    var timeNames = { MORNING: '朝', NOON: '昼', EVENING: '夕方', NIGHT: '夜' };
    var logText = player.name + ' が時間を進めました。【' + nextDay + '日目・' + timeNames[nextTime] + '】になりました。';
    addLog(ss, islandId, playerId, nextDay, logText);

    // 全員死亡判定 -> GAMEOVER
    if (allDead) {
      islandsSheet.getRange(island.row, 7).setValue('GAMEOVER');
      addLog(ss, islandId, 'SYSTEM', nextDay, '無情にも生存者は全員息絶えてしまった……GAME OVER');
      return {
        success: true,
        islandStatus: 'GAMEOVER',
        message: '全滅しました……'
      };
    }

    // 30日生存達成判定
    if (nextDay >= 30) {
      // 30日生存目標達成！
      recordAchievement(ss, playerId, '30日間過酷な島で生存した');
      addLog(ss, islandId, 'SYSTEM', nextDay, '生存30日を達成！救援船の目撃情報が届いた！');
    }

    // ランダムイベント判定（夕方・夜の進行時に25%の確率で発生）
    if (Math.random() < 0.25) {
      triggerRandomEvent(ss, islandId, nextDay, hasCabin);
    }

    island.day = nextDay;
    island.time_period = nextTime;
    var myLatestPlayer = authenticatePlayer(ss, playerId, token);
    SpreadsheetApp.flush();

    return {
      success: true,
      day: nextDay,
      timePeriod: nextTime,
      island: island,
      player: myLatestPlayer,
      islandStatus: 'ACTIVE',
      message: logText
    };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ランダムイベントの実行
 */
function triggerRandomEvent(ss, islandId, day, hasCabin) {
  var eventsSheet = ss.getSheetByName(SHEETS.EVENTS);
  var events = eventsSheet.getDataRange().getValues();
  if (events.length <= 1) return;

  var randIdx = 1 + Math.floor(Math.random() * (events.length - 1));
  var ev = events[randIdx];
  var evId = ev[0];
  var evName = ev[1];
  var evDesc = ev[2];

  var eventMsg = '【島内イベント発生】' + evName + ' : ' + evDesc;

  if (evId === 'STORM') {
    if (hasCabin) {
      eventMsg += '（木の小屋のおかげで被害を防いだ！）';
    } else {
      eventMsg += '（全員が暴風雨の被害を受けた）';
      // メンバーのHP微小減少
      damageAllActiveMembers(ss, islandId, 10);
    }
  } else if (evId === 'FIND_FOOD') {
    addToSharedInventory(ss, islandId, 'berry', 5);
    addToSharedInventory(ss, islandId, 'water', 3);
    eventMsg += '（共有倉庫に木の実×5, 水×3が追加されました）';
  } else if (evId === 'TREASURE_CHEST') {
    addToSharedInventory(ss, islandId, 'iron_ore', 2);
    addToSharedInventory(ss, islandId, 'cloth', 2);
    addToSharedInventory(ss, islandId, 'rope', 2);
    eventMsg += '（共有倉庫に鉄鉱石×2, 布×2, ロープ×2が追加されました）';
  }

  addLog(ss, islandId, 'EVENT', day, eventMsg);
}

function damageAllActiveMembers(ss, islandId, amount) {
  var membersSheet = ss.getSheetByName(SHEETS.MEMBERS);
  var mData = membersSheet.getDataRange().getValues();
  var playersSheet = ss.getSheetByName(SHEETS.PLAYERS);
  var pData = playersSheet.getDataRange().getValues();

  for (var m = 1; m < mData.length; m++) {
    if (mData[m][0] === islandId && mData[m][3] === 'ACTIVE') {
      var pId = mData[m][1];
      for (var p = 1; p < pData.length; p++) {
        if (pData[p][0] === pId && pData[p][7] === 'ALIVE') {
          var hp = Number(pData[p][3]);
          var nextHp = Math.max(0, hp - amount);
          playersSheet.getRange(p + 1, 4).setValue(nextHp);
          if (nextHp === 0) {
            playersSheet.getRange(p + 1, 8).setValue('DEAD');
          }
        }
      }
    }
  }
}

/**
 * 共有倉庫へのアイテム追加ヘルパー
 */
function addToSharedInventory(ss, islandId, itemId, quantity) {
  var sharedSheet = ss.getSheetByName(SHEETS.SHARED_INVENTORY);
  var sData = sharedSheet.getDataRange().getValues();
  var foundRow = -1;
  for (var i = 1; i < sData.length; i++) {
    if (sData[i][0] === islandId && sData[i][1] === itemId) {
      foundRow = i + 1;
      break;
    }
  }
  if (foundRow > 0) {
    var cur = Number(sharedSheet.getRange(foundRow, 3).getValue());
    sharedSheet.getRange(foundRow, 3).setValue(cur + quantity);
  } else {
    sharedSheet.appendRow([islandId, itemId, quantity]);
  }
}

/**
 * 個人インベントリと共有倉庫間のアイテム移動
 * 同時実行時にも数量が壊れない排他制御
 */
function transferItem(playerId, token, islandId, action, itemId, quantity) {
  quantity = parseInt(quantity, 10);
  if (isNaN(quantity) || quantity <= 0) {
    return { success: false, message: '数量が不正です。' };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var player = authenticatePlayer(ss, playerId, token);
    verifyIslandMembership(ss, islandId, playerId);
    var island = getIslandData(ss, islandId);

    var invSheet = ss.getSheetByName(SHEETS.INVENTORY);
    var sharedSheet = ss.getSheetByName(SHEETS.SHARED_INVENTORY);

    var invData = invSheet.getDataRange().getValues();
    var sharedData = sharedSheet.getDataRange().getValues();

    if (action === 'deposit') {
      // 個人 -> 共有倉庫
      var pRow = -1;
      var pQty = 0;
      for (var p = 1; p < invData.length; p++) {
        if (invData[p][0] === playerId && invData[p][1] === itemId) {
          pRow = p + 1;
          pQty = Number(invData[p][2]);
          break;
        }
      }

      if (pRow === -1 || pQty < quantity) {
        return { success: false, message: '預けるアイテムの所持数が不足しています。' };
      }

      // 個人側を減算
      var remain = pQty - quantity;
      if (remain <= 0) {
        invSheet.deleteRow(pRow);
      } else {
        invSheet.getRange(pRow, 3).setValue(remain);
      }

      // 共有倉庫側を加算
      addToSharedInventory(ss, islandId, itemId, quantity);
      addLog(ss, islandId, playerId, island.day, player.name + ' が共有倉庫に [' + itemId + '×' + quantity + '] を預けました。');

    } else if (action === 'withdraw') {
      // 共有倉庫 -> 個人
      var sRow = -1;
      var sQty = 0;
      for (var s = 1; s < sharedData.length; s++) {
        if (sharedData[s][0] === islandId && sharedData[s][1] === itemId) {
          sRow = s + 1;
          sQty = Number(sharedData[s][2]);
          break;
        }
      }

      if (sRow === -1 || sQty < quantity) {
        return { success: false, message: '共有倉庫の在庫が不足しています。' };
      }

      // 共有側を減算
      var remainShared = sQty - quantity;
      if (remainShared <= 0) {
        sharedSheet.deleteRow(sRow);
      } else {
        sharedSheet.getRange(sRow, 3).setValue(remainShared);
      }

      // 個人側を加算
      var pRow2 = -1;
      for (var p2 = 1; p2 < invData.length; p2++) {
        if (invData[p2][0] === playerId && invData[p2][1] === itemId) {
          pRow2 = p2 + 1;
          break;
        }
      }
      if (pRow2 > 0) {
        var curP = Number(invSheet.getRange(pRow2, 3).getValue());
        invSheet.getRange(pRow2, 3).setValue(curP + quantity);
      } else {
        invSheet.appendRow([playerId, itemId, quantity]);
      }

      addLog(ss, islandId, playerId, island.day, player.name + ' が共有倉庫から [' + itemId + '×' + quantity + '] を引き出しました。');
    } else {
      return { success: false, message: '不正なアクションです。' };
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * プレイヤー間アイテム直接譲渡
 */
function giveItemToPlayer(playerId, token, islandId, targetPlayerId, itemId, quantity) {
  quantity = parseInt(quantity, 10);
  if (isNaN(quantity) || quantity <= 0) {
    return { success: false, message: '数量が不正です。' };
  }
  if (playerId === targetPlayerId) {
    return { success: false, message: '自分自身には渡せません。' };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var sender = authenticatePlayer(ss, playerId, token);
    verifyIslandMembership(ss, islandId, playerId);
    verifyIslandMembership(ss, islandId, targetPlayerId);
    var island = getIslandData(ss, islandId);

    // 相手の名前取得
    var pSheet = ss.getSheetByName(SHEETS.PLAYERS);
    var pData = pSheet.getDataRange().getValues();
    var targetName = '仲間';
    for (var i = 1; i < pData.length; i++) {
      if (pData[i][0] === targetPlayerId) {
        targetName = pData[i][1];
        break;
      }
    }

    var invSheet = ss.getSheetByName(SHEETS.INVENTORY);
    var invData = invSheet.getDataRange().getValues();

    var senderRow = -1;
    var senderQty = 0;
    var targetRow = -1;
    var targetQty = 0;

    for (var r = 1; r < invData.length; r++) {
      if (invData[r][0] === playerId && invData[r][1] === itemId) {
        senderRow = r + 1;
        senderQty = Number(invData[r][2]);
      }
      if (invData[r][0] === targetPlayerId && invData[r][1] === itemId) {
        targetRow = r + 1;
        targetQty = Number(invData[r][2]);
      }
    }

    if (senderRow === -1 || senderQty < quantity) {
      return { success: false, message: '所持数が不足しています。' };
    }

    // 差出人の減算
    if (senderQty - quantity <= 0) {
      invSheet.deleteRow(senderRow);
    } else {
      invSheet.getRange(senderRow, 3).setValue(senderQty - quantity);
    }

    // 受取人の加算
    if (targetRow > 0) {
      // 差出人の行削除により行がずれている可能性を考慮
      var actualTargetRow = (senderRow > 0 && senderRow < targetRow && (senderQty - quantity <= 0)) ? targetRow - 1 : targetRow;
      invSheet.getRange(actualTargetRow, 3).setValue(targetQty + quantity);
    } else {
      invSheet.appendRow([targetPlayerId, itemId, quantity]);
    }

    addLog(ss, islandId, playerId, island.day, sender.name + ' が ' + targetName + ' に [' + itemId + '×' + quantity + '] を手渡しました。');

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * アイテム使用（食事・水分補給・回復）
 */
function useItem(playerId, token, islandId, itemId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var player = authenticatePlayer(ss, playerId, token);
    verifyIslandMembership(ss, islandId, playerId);
    var island = getIslandData(ss, islandId);

    var invSheet = ss.getSheetByName(SHEETS.INVENTORY);
    var invData = invSheet.getDataRange().getValues();
    var row = -1;
    var qty = 0;
    for (var i = 1; i < invData.length; i++) {
      if (invData[i][0] === playerId && invData[i][1] === itemId) {
        row = i + 1;
        qty = Number(invData[i][2]);
        break;
      }
    }

    if (row === -1 || qty < 1) {
      return { success: false, message: 'アイテムを持っていません。' };
    }

    // アイテム効果の適用
    var isMedic = (player.role === 'MEDIC');
    var hpGain = 0;
    var hungerGain = 0;
    var thirstGain = 0;
    var staminaGain = 0;
    var effectDesc = '';

    if (itemId === 'berry') {
      hungerGain = 20;
      thirstGain = 5;
      effectDesc = '木の実を食べ、空腹が満たされた。';
    } else if (itemId === 'water') {
      thirstGain = 35;
      effectDesc = 'きれいな水を飲み、渇きが癒やされた。';
    } else if (itemId === 'fish') {
      hungerGain = 35;
      hpGain = 10;
      effectDesc = '魚を食べ、力が湧いてきた。';
    } else if (itemId === 'herb') {
      hpGain = isMedic ? 40 : 25; // MEDICは回復量大幅アップ
      effectDesc = '薬草で傷を手当てした。' + (isMedic ? '（メディック補正発動！）' : '');
    } else {
      return { success: false, message: 'このアイテムは使用できません。' };
    }

    // アイテム消費
    if (qty === 1) {
      invSheet.deleteRow(row);
    } else {
      invSheet.getRange(row, 3).setValue(qty - 1);
    }

    // ステータス反映
    var playersSheet = ss.getSheetByName(SHEETS.PLAYERS);
    var newHp = Math.min(100, player.hp + hpGain);
    var newHunger = Math.min(100, player.hunger + hungerGain);
    var newThirst = Math.min(100, player.thirst + thirstGain);
    var newStamina = Math.min(100, player.stamina + staminaGain);

    playersSheet.getRange(player.row, 4).setValue(newHp);
    playersSheet.getRange(player.row, 5).setValue(newHunger);
    playersSheet.getRange(player.row, 6).setValue(newThirst);
    playersSheet.getRange(player.row, 7).setValue(newStamina);

    addLog(ss, islandId, playerId, island.day, player.name + ' が [' + itemId + '] を使用しました。' + effectDesc);

    player.hp = newHp;
    player.hunger = newHunger;
    player.thirst = newThirst;
    player.stamina = newStamina;
    SpreadsheetApp.flush();

    return {
      success: true,
      player: player,
      message: effectDesc
    };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// STEP 5: クラフト、建築、役割システム
// ==========================================

/**
 * クラフト処理
 */
function craftItem(playerId, token, islandId, recipeId, useShared) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var player = authenticatePlayer(ss, playerId, token);
    verifyIslandMembership(ss, islandId, playerId);
    var island = getIslandData(ss, islandId);

    // レシピ取得
    var recSheet = ss.getSheetByName(SHEETS.RECIPES);
    var rData = recSheet.getDataRange().getValues();
    var recipe = null;
    for (var r = 1; r < rData.length; r++) {
      if (rData[r][0] === recipeId) {
        recipe = {
          recipe_id: rData[r][0],
          result_item_id: rData[r][1],
          wood: Number(rData[r][2]),
          stone: Number(rData[r][3]),
          other_material: rData[r][4] // e.g. "rope:1" or "herb:2"
        };
        break;
      }
    }

    if (!recipe) {
      return { success: false, message: '存在しないレシピです。' };
    }

    // CRAFTSMAN（職人）は素材消費20%軽減（端数切り捨て、最低1個）
    var isCraftsman = (player.role === 'CRAFTSMAN');
    var reqWood = recipe.wood;
    var reqStone = recipe.stone;
    if (isCraftsman) {
      if (reqWood > 1) reqWood = Math.max(1, Math.floor(reqWood * 0.8));
      if (reqStone > 1) reqStone = Math.max(1, Math.floor(reqStone * 0.8));
    }

    // その他素材のパース (e.g. "rope:2,cloth:1")
    var reqOthers = [];
    if (recipe.other_material && recipe.other_material.trim() !== '') {
      var parts = recipe.other_material.split(',');
      for (var p = 0; p < parts.length; p++) {
        var sub = parts[p].split(':');
        if (sub.length === 2) {
          var oQty = Number(sub[1]);
          if (isCraftsman && oQty > 1) oQty = Math.max(1, Math.floor(oQty * 0.8));
          reqOthers.push({ itemId: sub[0].trim(), quantity: oQty });
        }
      }
    }

    // 消費元（個人インベントリ または 共有倉庫）
    var targetSheet = useShared ? ss.getSheetByName(SHEETS.SHARED_INVENTORY) : ss.getSheetByName(SHEETS.INVENTORY);
    var tData = targetSheet.getDataRange().getValues();

    // 在庫チェック
    function getItemStock(itmId) {
      for (var i = 1; i < tData.length; i++) {
        var ownerMatch = useShared ? (tData[i][0] === islandId) : (tData[i][0] === playerId);
        if (ownerMatch && tData[i][1] === itmId) {
          return { row: i + 1, qty: Number(tData[i][2]) };
        }
      }
      return { row: -1, qty: 0 };
    }

    if (reqWood > 0 && getItemStock('wood').qty < reqWood) {
      return { success: false, message: '木材が不足しています（必要: ' + reqWood + '）。' };
    }
    if (reqStone > 0 && getItemStock('stone').qty < reqStone) {
      return { success: false, message: '石が不足しています（必要: ' + reqStone + '）。' };
    }
    for (var o = 0; o < reqOthers.length; o++) {
      if (getItemStock(reqOthers[o].itemId).qty < reqOthers[o].quantity) {
        return { success: false, message: reqOthers[o].itemId + ' が不足しています。' };
      }
    }

    // 素材消費処理
    function deductMaterial(itmId, amount) {
      if (amount <= 0) return;
      var curData = targetSheet.getDataRange().getValues();
      for (var i = 1; i < curData.length; i++) {
        var ownerMatch = useShared ? (curData[i][0] === islandId) : (curData[i][0] === playerId);
        if (ownerMatch && curData[i][1] === itmId) {
          var curQty = Number(curData[i][2]);
          if (curQty <= amount) {
            targetSheet.deleteRow(i + 1);
          } else {
            targetSheet.getRange(i + 1, 3).setValue(curQty - amount);
          }
          break;
        }
      }
    }

    if (reqWood > 0) deductMaterial('wood', reqWood);
    if (reqStone > 0) deductMaterial('stone', reqStone);
    for (var o2 = 0; o2 < reqOthers.length; o2++) {
      deductMaterial(reqOthers[o2].itemId, reqOthers[o2].quantity);
    }

    // 完成品を個人インベントリへ付与
    var invSheet = ss.getSheetByName(SHEETS.INVENTORY);
    var invData = invSheet.getDataRange().getValues();
    var resRow = -1;
    for (var ir = 1; ir < invData.length; ir++) {
      if (invData[ir][0] === playerId && invData[ir][1] === recipe.result_item_id) {
        resRow = ir + 1;
        break;
      }
    }
    if (resRow > 0) {
      var prevQty = Number(invSheet.getRange(resRow, 3).getValue());
      invSheet.getRange(resRow, 3).setValue(prevQty + 1);
    } else {
      invSheet.appendRow([playerId, recipe.result_item_id, 1]);
    }

    recordAchievement(ss, playerId, '初めてのクラフト成功');
    var logMsg = player.name + ' が [' + recipe.result_item_id + '] をクラフトしました！' + (isCraftsman ? '（職人効果で素材節約）' : '');
    addLog(ss, islandId, playerId, island.day, logMsg);

    return {
      success: true,
      resultItemId: recipe.result_item_id,
      message: logMsg
    };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 拠点建築システム
 */
function buildStructure(playerId, token, islandId, buildingType) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var player = authenticatePlayer(ss, playerId, token);
    verifyIslandMembership(ss, islandId, playerId);
    var island = getIslandData(ss, islandId);

    // 建築物コスト定義
    var buildingCosts = {
      campfire: { name: '焚き火', wood: 5, stone: 3, other: [], maxLevel: 1 },
      tent: { name: 'テント', wood: 8, stone: 0, other: [{ itemId: 'cloth', quantity: 2 }, { itemId: 'rope', quantity: 2 }], maxLevel: 1 },
      cabin: { name: '木の小屋', wood: 20, stone: 10, other: [{ itemId: 'rope', quantity: 4 }], maxLevel: 1 },
      water_collector: { name: '貯水設備', wood: 10, stone: 5, other: [{ itemId: 'cloth', quantity: 2 }], maxLevel: 1 },
      raft: { name: '脱出用いかだ', wood: 15, stone: 5, other: [{ itemId: 'rope', quantity: 5 }], maxLevel: 3 }
    };

    var bDef = buildingCosts[buildingType];
    if (!bDef) {
      return { success: false, message: '無効な建築物タイプです。' };
    }

    // 現在の建築状況確認
    var buildSheet = ss.getSheetByName(SHEETS.BUILDINGS);
    var bData = buildSheet.getDataRange().getValues();
    var existingRow = -1;
    var currentLevel = 0;

    for (var b = 1; b < bData.length; b++) {
      if (bData[b][1] === islandId && bData[b][2] === buildingType) {
        existingRow = b + 1;
        currentLevel = Number(bData[b][3]);
        break;
      }
    }

    if (currentLevel >= bDef.maxLevel) {
      return { success: false, message: 'この設備は既に最大レベルです。' };
    }

    var nextLevel = currentLevel + 1;
    var isCraftsman = (player.role === 'CRAFTSMAN');

    // レベルに応じたコスト倍率
    var multiplier = nextLevel;
    var reqWood = Math.floor(bDef.wood * multiplier * (isCraftsman ? 0.8 : 1.0));
    var reqStone = Math.floor(bDef.stone * multiplier * (isCraftsman ? 0.8 : 1.0));

    // 建築素材は共有倉庫から優先消費
    var sharedSheet = ss.getSheetByName(SHEETS.SHARED_INVENTORY);
    var sData = sharedSheet.getDataRange().getValues();

    function getSharedStock(itmId) {
      for (var s = 1; s < sData.length; s++) {
        if (sData[s][0] === islandId && sData[s][1] === itmId) {
          return { row: s + 1, qty: Number(sData[s][2]) };
        }
      }
      return { row: -1, qty: 0 };
    }

    if (reqWood > 0 && getSharedStock('wood').qty < reqWood) {
      return { success: false, message: '共有倉庫の木材が不足しています（必要: ' + reqWood + '）。' };
    }
    if (reqStone > 0 && getSharedStock('stone').qty < reqStone) {
      return { success: false, message: '共有倉庫の石が不足しています（必要: ' + reqStone + '）。' };
    }

    // その他素材のチェック
    for (var ot = 0; ot < bDef.other.length; ot++) {
      var reqOtherQty = Math.floor(bDef.other[ot].quantity * multiplier * (isCraftsman ? 0.8 : 1.0));
      if (getSharedStock(bDef.other[ot].itemId).qty < reqOtherQty) {
        return { success: false, message: '共有倉庫の ' + bDef.other[ot].itemId + ' が不足しています。' };
      }
    }

    // 素材減算
    function deductShared(itmId, amount) {
      if (amount <= 0) return;
      var curS = sharedSheet.getDataRange().getValues();
      for (var i = 1; i < curS.length; i++) {
        if (curS[i][0] === islandId && curS[i][1] === itmId) {
          var curQty = Number(curS[i][2]);
          if (curQty <= amount) {
            sharedSheet.deleteRow(i + 1);
          } else {
            sharedSheet.getRange(i + 1, 3).setValue(curQty - amount);
          }
          break;
        }
      }
    }

    if (reqWood > 0) deductShared('wood', reqWood);
    if (reqStone > 0) deductShared('stone', reqStone);
    for (var ot2 = 0; ot2 < bDef.other.length; ot2++) {
      var rQty = Math.floor(bDef.other[ot2].quantity * multiplier * (isCraftsman ? 0.8 : 1.0));
      deductShared(bDef.other[ot2].itemId, rQty);
    }

    // 建築物データの記録
    var now = getNowString();
    if (existingRow > 0) {
      buildSheet.getRange(existingRow, 4).setValue(nextLevel);
      buildSheet.getRange(existingRow, 5).setValue(player.name);
      buildSheet.getRange(existingRow, 6).setValue(now);
    } else {
      buildSheet.appendRow([generateUUID(), islandId, buildingType, nextLevel, player.name, now]);
    }

    var logMsg = player.name + ' が【' + bDef.name + ' (Lv.' + nextLevel + ')】を建築しました！';
    addLog(ss, islandId, playerId, island.day, logMsg);
    recordAchievement(ss, playerId, '拠点の建築主');

    return {
      success: true,
      buildingType: buildingType,
      level: nextLevel,
      message: logMsg
    };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 仲間救助アクション（MEDICまたは薬草所持者）
 */
function rescuePlayer(playerId, token, islandId, targetPlayerId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var rescuer = authenticatePlayer(ss, playerId, token);
    verifyIslandMembership(ss, islandId, playerId);
    verifyIslandMembership(ss, islandId, targetPlayerId);
    var island = getIslandData(ss, islandId);

    if (rescuer.status === 'DEAD') {
      return { success: false, message: '自身が死亡している状態では救助できません。' };
    }

    var isMedic = (rescuer.role === 'MEDIC');
    var playersSheet = ss.getSheetByName(SHEETS.PLAYERS);
    var pData = playersSheet.getDataRange().getValues();

    var targetRow = -1;
    var targetName = '';
    var targetStatus = '';

    for (var p = 1; p < pData.length; p++) {
      if (pData[p][0] === targetPlayerId) {
        targetRow = p + 1;
        targetName = pData[p][1];
        targetStatus = pData[p][7];
        break;
      }
    }

    if (targetStatus !== 'DEAD') {
      return { success: false, message: 'そのプレイヤーは生存しています。' };
    }

    // 薬草(herb)または真水(water)を個人インベントリから消費
    var invSheet = ss.getSheetByName(SHEETS.INVENTORY);
    var invData = invSheet.getDataRange().getValues();
    var herbRow = -1;
    for (var i = 1; i < invData.length; i++) {
      if (invData[i][0] === playerId && (invData[i][1] === 'herb' || invData[i][1] === 'water')) {
        herbRow = i + 1;
        break;
      }
    }

    if (herbRow === -1 && !isMedic) {
      return { success: false, message: '救助には薬草または真水が必要です。' };
    }

    // 救助物資消費（メディックは物資がなくてもスタミナ30消費で蘇生可能）
    if (herbRow > 0) {
      var hQty = Number(invSheet.getRange(herbRow, 3).getValue());
      if (hQty <= 1) {
        invSheet.deleteRow(herbRow);
      } else {
        invSheet.getRange(herbRow, 3).setValue(hQty - 1);
      }
    } else {
      if (rescuer.stamina < 30) {
        return { success: false, message: '救助に必要なスタミナが不足しています。' };
      }
      playersSheet.getRange(rescuer.row, 7).setValue(rescuer.stamina - 30);
    }

    // 蘇生処理
    var reviveHp = isMedic ? 40 : 20;
    playersSheet.getRange(targetRow, 4).setValue(reviveHp);
    playersSheet.getRange(targetRow, 5).setValue(30); // 飢餓微小回復
    playersSheet.getRange(targetRow, 6).setValue(30); // 渇き微小回復
    playersSheet.getRange(targetRow, 8).setValue('ALIVE');

    var logMsg = rescuer.name + ' が瀕死の ' + targetName + ' を救助・蘇生させました！';
    addLog(ss, islandId, playerId, island.day, logMsg);
    recordAchievement(ss, playerId, '救命の英雄');

    return { success: true, message: logMsg };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// STEP 6: チャット、実績システム
// ==========================================

function sendMessage(playerId, token, islandId, message) {
  if (!message || message.trim() === '') {
    return { success: false, message: 'メッセージを入力してください。' };
  }
  if (message.length > 200) {
    return { success: false, message: 'メッセージは200文字以内にしてください。' };
  }

  try {
    var ss = getSpreadsheet();
    var player = authenticatePlayer(ss, playerId, token);
    verifyIslandMembership(ss, islandId, playerId);

    var cleanMessage = message.trim()
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');

    var msgSheet = ss.getSheetByName(SHEETS.MESSAGES);
    msgSheet.appendRow([
      generateUUID(),
      islandId,
      playerId,
      player.name,
      cleanMessage,
      getNowString()
    ]);

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function recordAchievement(ss, playerId, achievementName) {
  try {
    var achSheet = ss.getSheetByName(SHEETS.ACHIEVEMENTS);
    var aData = achSheet.getDataRange().getValues();
    for (var i = 1; i < aData.length; i++) {
      if (aData[i][1] === playerId && aData[i][2] === achievementName) {
        return; // 重複記録防止
      }
    }
    achSheet.appendRow([generateUUID(), playerId, achievementName, getNowString()]);
  } catch (e) {}
}

function getAchievements(playerId, token) {
  try {
    var ss = getSpreadsheet();
    authenticatePlayer(ss, playerId, token);
    var achSheet = ss.getSheetByName(SHEETS.ACHIEVEMENTS);
    var aData = achSheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < aData.length; i++) {
      if (aData[i][1] === playerId) {
        list.push({
          achievementId: aData[i][0],
          name: aData[i][2],
          unlockedAt: aData[i][3]
        });
      }
    }
    return { success: true, achievements: list };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ==========================================
// STEP 7: 勝利（脱出）処理
// ==========================================

function escapeIsland(playerId, token, islandId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var player = authenticatePlayer(ss, playerId, token);
    verifyIslandMembership(ss, islandId, playerId);
    var island = getIslandData(ss, islandId);

    if (island.island_status !== 'ACTIVE') {
      return { success: false, message: '既にゲームは終了しています。' };
    }

    // 1. いかだがレベル3（完成）か確認
    var buildSheet = ss.getSheetByName(SHEETS.BUILDINGS);
    var bData = buildSheet.getDataRange().getValues();
    var raftLevel = 0;
    for (var b = 1; b < bData.length; b++) {
      if (bData[b][1] === islandId && bData[b][2] === 'raft') {
        raftLevel = Number(bData[b][3]);
        break;
      }
    }

    if (raftLevel < 3) {
      return { success: false, message: 'いかだが完成していません（必要: いかだ Lv.3）。クラフトと建築を進めてください。' };
    }

    // 2. 脱出用物資の確認（共有倉庫に 水×10、食料(魚 or 木の実)×10）
    var sharedSheet = ss.getSheetByName(SHEETS.SHARED_INVENTORY);
    var sData = sharedSheet.getDataRange().getValues();
    var waterCount = 0;
    var foodCount = 0;

    for (var s = 1; s < sData.length; s++) {
      if (sData[s][0] === islandId) {
        var itm = sData[s][1];
        var q = Number(sData[s][2]);
        if (itm === 'water') waterCount += q;
        if (itm === 'berry' || itm === 'fish') foodCount += q;
      }
    }

    if (waterCount < 10) {
      return { success: false, message: '脱出航海用の水が不足しています（共有倉庫に水×10必要。現在: ' + waterCount + '個）。' };
    }
    if (foodCount < 10) {
      return { success: false, message: '脱出航海用の食料が不足しています（共有倉庫に木の実または魚×10必要。現在: ' + foodCount + '個）。' };
    }

    // 勝利フラグ確定！
    var islandsSheet = ss.getSheetByName(SHEETS.ISLANDS);
    islandsSheet.getRange(island.row, 7).setValue('CLEARED');
    islandsSheet.getRange(island.row, 9).setValue(getNowString());

    var winLog = '【脱出成功！】生存者たちは完成したいかだに乗り込み、仲間と共に未知なる大海原へと船出しました！全員島からの脱出に成功！';
    addLog(ss, islandId, playerId, island.day, winLog);

    // 全メンバーに実績付与
    var membersSheet = ss.getSheetByName(SHEETS.MEMBERS);
    var mData = membersSheet.getDataRange().getValues();
    for (var m = 1; m < mData.length; m++) {
      if (mData[m][0] === islandId && mData[m][3] === 'ACTIVE') {
        recordAchievement(ss, mData[m][1], '無人島からの生還者（ゲームクリア）');
      }
    }

    return {
      success: true,
      islandStatus: 'CLEARED',
      message: winLog
    };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}
