/**
 * PowerPro - 選手保存を最初の空行に、打順・ポジション設定をバグ修正
 * 変化球拡張・投手左右・打者左右・変化球割合設定を追加
 */

const DATA_SHEET_PROP = 'POWERPRO_DATA_SHEET_ID';

function doGet(e) {
  ensureDataSheets();
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('PowerPro Web (Apps Script)');
}

function getOrCreateSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(DATA_SHEET_PROP);
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) {}
  }
  const ss = SpreadsheetApp.create('PowerPro_Data');
  props.setProperty(DATA_SHEET_PROP, ss.getId());
  return ss;
}

function ensureDataSheets() {
  const ss = getOrCreateSpreadsheet();
  const required = [
    {name:'Players', headers:[
      'id','teamId','teamRole','name','isPitcher','handedness',
      'meet','power','speed','defense','arm','catching',
      'armAngle','avgSpeed','stamina','control','pitches',
      'position','specialAbilities',
      'PA','AB','H','HR','BB','K','TB',
      'outsRecorded','earnedRuns','pitcherK',
      'created_at','updated_at'
    ]},
    {name:'Teams', headers:['id','name','playerIds','created_at','updated_at']},
    {name:'PitchTypes', headers:[
      'id','name','japaneseName','baseSpeed',
      'movementX','movementY','spinRate','spinAxis',
      'category','description','created_at','updated_at'
    ]},
    {name:'PitchRatios', headers:[
      'id','pitcherId','pitchTypeId','ratio','created_at','updated_at'
    ]},
    {name:'Meta', headers:['key','value']},
    {name:'Abilities', headers:['id','name','category','description']},
    {name:'Matches', headers:['matchId','homeTeamId','awayTeamId','date','summaryJson']}
  ];

  required.forEach(function(def){
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      sheet.getRange(1,1,1,def.headers.length).setValues([def.headers]);
    } else {
      const existing = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      if (existing.length < def.headers.length) {
        sheet.getRange(1,existing.length+1,1,def.headers.length-existing.length).setValues([def.headers.slice(existing.length)]);
      }
    }
  });

  const ptSheet = ss.getSheetByName('PitchTypes');
  if (ptSheet.getLastRow() <= 1) {
    initializePitchTypes();
  }

  const abilSheet = ss.getSheetByName('Abilities');
  if (abilSheet.getLastRow() <= 1) {
    const abilities = [
      ['powerhitter','パワーヒッター','batter','ホームラン狙いが多い'],
      ['averagehitter','アベレージヒッター','batter','ヒット狙いが多い'],
      ['starter_yes','初級〇','batter','初級に積極的'],
      ['sticky_fouls','粘りうち','batter','2ストライクからファールで粘りがち'],
      ['k_prone','三振','batter','2ストライクからミートが下がる'],
      ['doubleplay_prone','併殺','batter','1塁ランナーがいると三振が多くなる'],
      ['steal_good','盗塁〇','batter','盗塁が上手い'],
      ['steal_bad','盗塁×','batter','盗塁が下手'],
      ['opposite_hit','流し打ち','batter','流し方向が多い'],
      ['pull_hit','引っ張り','batter','引っ張りが多い'],
      ['nobi','ノビ','pitcher','ストレートの伸びが良い'],
      ['kire','キレ','pitcher','変化球の曲がるタイミングが遅くなる'],
      ['quick_good','クイック〇','pitcher','盗塁されにくい'],
      ['one_shot','一発','pitcher','ランナーがいるときの失投が真ん中に行きやすい'],
      ['k_boost','奪三振','pitcher','2ストライクに追い込むとコントロールが上がる'],
      ['walk_prone','四球','pitcher','3ボールになるとコントロールが悪くなる']
    ];
    abilSheet.getRange(2,1,abilities.length,4).setValues(abilities);
  }

  const teamsSheet = ss.getSheetByName('Teams');
  if (teamsSheet.getLastRow() <= 1) {
    createPracticeTeams();
  }
}

function initializePitchTypes() {
  const ss = getOrCreateSpreadsheet();
  const ptSheet = ss.getSheetByName('PitchTypes');
  const now = (new Date()).toISOString();
  
  const pitchTypes = [
    // ストレート系
    ['fastball', 'ストレート', 'ストレート', 145, 0, 0, 2200, 180, 'fastball', '基本となる直球。最速', now, now],
    ['riseball', 'ライズボール', 'ライズボール', 138, 0, 8, 2400, 180, 'fastball', '上昇する直球。難しい', now, now],
    ['cutter', 'カット', 'カット', 140, 12, -2, 2400, 60, 'fastball', 'わずかな横ムーブ。速さとの両立', now, now],
    
    // 横変化系
    ['slider', 'スライダー', 'スライダー', 135, 18, -8, 2600, 45, 'breaker', '横に鋭く曲がる。右打者に有効', now, now],
    ['slurve', 'スラーブ', 'スラーブ', 132, 22, -12, 2500, 30, 'breaker', 'スライダーとカーブの中間。軌道が複雑', now, now],
    ['screwball', 'スクリューボール', 'スクリューボール', 128, -15, -6, 2300, 315, 'breaker', '左投手の得意な変化球。左打者対策', now, now],
    
    // 縦変化系
    ['curveball', 'カーブ', 'カーブ', 125, 8, -25, 1800, 90, 'breaker', '落差が大きい変化球。コース指定が重要', now, now],
    ['fork', 'フォーク', 'フォーク', 130, -2, -30, 1500, 270, 'offspeed', '大きく落ちる。打者のタイミングを狂わす', now, now],
    ['splitter', 'スプリット', 'スプリット', 132, 0, -22, 1700, 270, 'offspeed', 'フォークより落ちが小さい。制御しやすい', now, now],
    
    // チェンジアップ系
    ['changeup', 'チェンジアップ', 'チェンジアップ', 115, 6, -8, 1200, 190, 'offspeed', 'ストレートに見えて遅い。タイミング狂わし', now, now],
    ['palmball', 'パームボール', 'パームボール', 110, 4, -10, 500, 180, 'offspeed', 'チェンジアップの一種。速度落ちが大きい', now, now],
    ['knuckleball', 'ナックル', 'ナックル', 90, 15, -8, 0, 0, 'offspeed', 'ほぼ無回転。変化が予測不可能', now, now]
  ];
  
  ptSheet.getRange(2, 1, pitchTypes.length, 12).setValues(pitchTypes);
}

/* --- Players CRUD --- */
function listPlayers() {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName('Players');
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  return rows.slice(1).map(function(r){
    const obj = {};
    headers.forEach(function(h,i){ obj[h]=r[i]; });
    if (!obj.specialAbilities) obj.specialAbilities = [];
    else if (typeof obj.specialAbilities === 'string') obj.specialAbilities = obj.specialAbilities ? obj.specialAbilities.split(',') : [];
    obj.isPitcher = (obj.isPitcher === true || obj.isPitcher === 'true' || obj.isPitcher === 1);
    obj.handedness = obj.handedness || 'right';
    ['meet','power','speed','defense','arm','catching','armAngle','avgSpeed','stamina','control',
     'PA','AB','H','HR','BB','K','TB','outsRecorded','earnedRuns','pitcherK'].forEach(function(k){
      if (obj[k] === '') obj[k] = 0;
      if (obj[k] != null && obj[k] !== '') obj[k] = Number(obj[k]);
    });
    return obj;
  });
}

function savePlayer(player) {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName('Players');
  const rows = sheet.getDataRange().getValues();
  if (!player.id) player.id = Utilities.getUuid();
  player.updated_at = (new Date()).toISOString();
  if (!player.created_at) player.created_at = player.updated_at;

  const pitcherPositions = ['先発','中継ぎ','抑え'];
  player.isPitcher = pitcherPositions.indexOf(player.position) >= 0;
  player.handedness = player.handedness || 'right';

  const specials = Array.isArray(player.specialAbilities) ? player.specialAbilities.join(',') : (player.specialAbilities || '');
  player.PA = player.PA || 0; player.AB = player.AB || 0; player.H = player.H || 0; player.HR = player.HR || 0;
  player.BB = player.BB || 0; player.K = player.K || 0; player.TB = player.TB || 0;
  player.outsRecorded = player.outsRecorded || 0; player.earnedRuns = player.earnedRuns || 0; player.pitcherK = player.pitcherK || 0;

  let foundRow = null;
  for (let r=1;r<rows.length;r++){
    if (rows[r][0] === player.id) { foundRow = r+1; break; }
  }

  const record = [
    player.id||'',
    player.teamId||'',
    player.teamRole||'',
    player.name||'',
    player.isPitcher || false,
    player.handedness || 'right',
    player.meet||0,
    player.power||0,
    player.speed||0,
    player.defense||0,
    player.arm||0,
    player.catching||0,
    player.armAngle||0,
    player.avgSpeed||0,
    player.stamina||0,
    player.control||0,
    player.position || '',
    specials,
    player.PA, player.AB, player.H, player.HR, player.BB, player.K, player.TB,
    player.outsRecorded, player.earnedRuns, player.pitcherK,
    player.created_at,
    player.updated_at
  ];

  if (foundRow) {
    sheet.getRange(foundRow,1,1,record.length).setValues([record]);
  } else {
    let insertRow = sheet.getLastRow() + 1;
    for (let r=2;r<=sheet.getLastRow();r++){
      const val = sheet.getRange(r,1).getValue();
      if (!val || val === '') {
        insertRow = r;
        break;
      }
    }
    sheet.getRange(insertRow,1,1,record.length).setValues([record]);
  }

  return player;
}

function deletePlayer(playerId) {
  if (!playerId) return {deleted:false, reason:'no id'};
  const ss = getOrCreateSpreadsheet();
  const players = ss.getSheetByName('Players');
  const rows = players.getDataRange().getValues();
  if (rows.length <= 1) return {deleted:false, reason:'no players'};
  const headers = rows[0];
  const idIndex = headers.indexOf('id');
  let foundRow = null;
  for (let r=1;r<rows.length;r++){
    if (rows[r][idIndex] === playerId) { foundRow = r+1; break; }
  }
  if (!foundRow) return {deleted:false, reason:'not found'};
  players.deleteRow(foundRow);

  const teams = ss.getSheetByName('Teams');
  const tRows = teams.getDataRange().getValues();
  if (tRows.length > 1) {
    const tHeaders = tRows[0];
    const pidIdx = tHeaders.indexOf('playerIds');
    for (let tr=1; tr<tRows.length; tr++){
      const val = tRows[tr][pidIdx] || '';
      const arr = val ? String(val).split(',').filter(s=>s) : [];
      const newArr = arr.filter(x => x !== playerId);
      if (newArr.length !== arr.length) {
        teams.getRange(tr+1, pidIdx+1).setValue(newArr.join(','));
      }
    }
  }
  return {deleted:true};
}

function deleteTeam(teamId) {
  if (!teamId) return {deleted:false, reason:'no id'};
  const ss = getOrCreateSpreadsheet();
  const teams = ss.getSheetByName('Teams');
  const tRows = teams.getDataRange().getValues();
  if (tRows.length <= 1) return {deleted:false, reason:'no teams'};
  const tHeaders = tRows[0];
  const idIdx = tHeaders.indexOf('id');
  let foundRow = null;
  for (let r=1; r<tRows.length; r++){
    if (tRows[r][idIdx] === teamId) { foundRow = r+1; break; }
  }
  if (!foundRow) return {deleted:false, reason:'team not found'};

  teams.deleteRow(foundRow);

  const players = ss.getSheetByName('Players');
  const pRows = players.getDataRange().getValues();
  if (pRows.length > 1) {
    const pHeaders = pRows[0];
    const teamIdIdx = pHeaders.indexOf('teamId');
    for (let pr=1; pr<pRows.length; pr++){
      if (pRows[pr][teamIdIdx] === teamId) {
        players.getRange(pr+1, teamIdIdx+1).setValue('');
      }
    }
  }

  return {deleted:true};
}

/* --- Teams --- */
function listTeams() {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName('Teams');
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  return rows.slice(1).map(function(r){
    const obj = {};
    headers.forEach(function(h,i){ obj[h]=r[i]; });
    obj.playerIds = obj.playerIds ? (typeof obj.playerIds === 'string' ? (obj.playerIds ? obj.playerIds.split(',') : []) : obj.playerIds) : [];
    return obj;
  });
}

function saveTeam(team) {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName('Teams');
  const rows = sheet.getDataRange().getValues();
  if (!team.id) team.id = Utilities.getUuid();
  team.updated_at = (new Date()).toISOString();
  if (!team.created_at) team.created_at = team.updated_at;
  const playerIdsStr = (Array.isArray(team.playerIds) ? team.playerIds.join(',') : (team.playerIds || ''));
  const record = [team.id, team.name||('Team-'+team.id.slice(0,4)), playerIdsStr, team.created_at, team.updated_at];
  let foundRow = null;
  for (let r=1;r<rows.length;r++){
    if (rows[r][0] === team.id) { foundRow = r+1; break; }
  }
  if (foundRow) sheet.getRange(foundRow,1,1,record.length).setValues([record]); else sheet.appendRow(record);
  return team;
}

function assignPlayersToTeam(teamId, playerIds) {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName('Players');
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return {updated:0};
  const headers = rows[0];
  const idIndex = headers.indexOf('id');
  const teamIdIndex = headers.indexOf('teamId');
  const idToRow = {};
  for (let r=1;r<rows.length;r++){
    idToRow[rows[r][idIndex]] = r+1;
  }
  let updated = 0;
  (playerIds || []).forEach(pid => {
    if (!pid) return;
    const rowIdx = idToRow[pid];
    if (!rowIdx) return;
    sheet.getRange(rowIdx, teamIdIndex+1).setValue(teamId);
    updated++;
  });
  return {updated: updated};
}

/* --- PitchTypes --- */
function listPitchTypes() {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName('PitchTypes');
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  return rows.slice(1).map(function(r){
    const obj = {};
    headers.forEach(function(h,i){ obj[h]=r[i]; });
    obj.movementX = Number(obj.movementX || 0);
    obj.movementY = Number(obj.movementY || 0);
    obj.baseSpeed = Number(obj.baseSpeed || 130);
    obj.spinRate = Number(obj.spinRate || 0);
    obj.spinAxis = Number(obj.spinAxis || 0);
    return obj;
  });
}

function savePitchType(pt) {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName('PitchTypes');
  const rows = sheet.getDataRange().getValues();
  if (!pt.id) pt.id = Utilities.getUuid();
  pt.updated_at = (new Date()).toISOString();
  if (!pt.created_at) pt.created_at = pt.updated_at;
  const record = [
    pt.id,
    pt.name||('球種-'+pt.id.slice(0,4)),
    pt.japaneseName || pt.name || '',
    pt.baseSpeed||130,
    pt.movementX||0,
    pt.movementY||0,
    pt.spinRate||0,
    pt.spinAxis||0,
    pt.category || 'fastball',
    pt.description || '',
    pt.created_at,
    pt.updated_at
  ];
  let foundRow = null;
  for (let r=1;r<rows.length;r++){
    if (rows[r][0] === pt.id) { foundRow = r+1; break; }
  }
  if (foundRow) sheet.getRange(foundRow,1,1,record.length).setValues([record]); else sheet.appendRow(record);
  return pt;
}

/* --- PitchRatios (投手の球種割合) --- */
function listPitchRatios(pitcherId) {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName('PitchRatios');
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  return rows.slice(1)
    .filter(r => r[1] === pitcherId)
    .map(function(r){
      const obj = {};
      headers.forEach(function(h,i){ obj[h]=r[i]; });
      obj.ratio = Number(obj.ratio || 0);
      return obj;
    });
}

function savePitchRatios(pitcherId, ratioList) {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName('PitchRatios');
  const rows = sheet.getDataRange().getValues();
  const now = (new Date()).toISOString();
  
  // 既存のこの投手の球種割合を削除
  for (let r = rows.length - 1; r >= 1; r--) {
    if (rows[r][1] === pitcherId) {
      sheet.deleteRow(r + 1);
      rows.splice(r, 1);
    }
  }
  
  // 新規追加
  ratioList.forEach(function(item) {
    const record = [
      Utilities.getUuid(),
      pitcherId,
      item.pitchTypeId,
      item.ratio || 0,
      now,
      now
    ];
    sheet.appendRow(record);
  });
  
  return {saved: true};
}

/* --- Abilities --- */
function listAbilities() {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName('Abilities');
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  return rows.slice(1).map(function(r){
    const obj = {};
    headers.forEach(function(h,i){ obj[h]=r[i]; });
    return obj;
  });
}

function simulateMatch(homeTeamId, awayTeamId, innings) {
  throw new Error('simulateMatch: 実装をここに置いてください（前回の完全版を貼り付けて下さい）');
}

function getGameSetup() {
  return {
    players: listPlayers(),
    teams: listTeams(),
    pitchTypes: listPitchTypes(),
    abilities: listAbilities(),
    config: { roster: {batters:12, starters:3, middle:5, closers:2}, strikeZoneGrid: {cols:9, rows:9} }
  };
}

function getPitchRatiosForPitcher(pitcherId) {
  return listPitchRatios(pitcherId);
}

function createPracticeTeams() {
  const ss = getOrCreateSpreadsheet();
  const teamAId = Utilities.getUuid();
  const teamBId = Utilities.getUuid();
  const teamAPlayerIds = [];
  const teamBPlayerIds = [];

  function makeBatter(name, teamId, pos, handedness) {
    const p = {
      id: Utilities.getUuid(),
      teamId: teamId,
      teamRole: '野手',
      name: name,
      isPitcher: false,
      handedness: handedness || 'right',
      meet:50,power:50,speed:50,defense:50,arm:50,catching:50,
      armAngle:45,avgSpeed:140,stamina:50,control:50,
      position: pos || '', specialAbilities:[],
      PA:0,AB:0,H:0,HR:0,BB:0,K:0,TB:0,
      outsRecorded:0,earnedRuns:0,pitcherK:0
    };
    savePlayer(p);
    return p.id;
  }
  function makePitcher(name, teamId, role, handedness) {
    const p = {
      id: Utilities.getUuid(),
      teamId: teamId,
      teamRole: '投手',
      name: name,
      isPitcher: true,
      handedness: handedness || 'right',
      meet:50,power:50,speed:50,defense:50,arm:50,catching:50,
      armAngle:45,avgSpeed:140,stamina:50,control:50,
      position: role || '', specialAbilities:[],
      PA:0,AB:0,H:0,HR:0,BB:0,K:0,TB:0,
      outsRecorded:0,earnedRuns:0,pitcherK:0
    };
    savePlayer(p);
    return p.id;
  }

  const positions = ['1B','2B','3B','SS','LF','CF','RF','C','DH','PH','LF2','RF2'];
  for (let i=1;i<=12;i++){
    const hand = (i % 3 === 0) ? 'left' : 'right';
    teamAPlayerIds.push(makeBatter('A選手'+i, teamAId, positions[(i-1)%positions.length], hand));
    teamBPlayerIds.push(makeBatter('B選手'+i, teamBId, positions[(i-1)%positions.length], hand));
  }
  for (let i=1;i<=10;i++){
    let role = (i<=3) ? '先発' : (i<=8 ? '中継ぎ' : '抑え');
    const hand = (i % 2 === 0) ? 'left' : 'right';
    teamAPlayerIds.push(makePitcher('A投手'+i, teamAId, role, hand));
    teamBPlayerIds.push(makePitcher('B投手'+i, teamBId, role, hand));
  }

  saveTeam({id:teamAId, name:'練習チームA', playerIds: teamAPlayerIds});
  saveTeam({id:teamBId, name:'練習チームB', playerIds: teamBPlayerIds});

  assignPlayersToTeam(teamAId, teamAPlayerIds);
  assignPlayersToTeam(teamBId, teamBPlayerIds);

  return {teamAId, teamBId};
}
