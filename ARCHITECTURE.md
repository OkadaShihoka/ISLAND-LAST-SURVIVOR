# ISLAND: LAST SURVIVOR - アーキテクチャ＆詳細設計書

## 1. ファイル構成
- `Code.gs` : Google Apps Script サーバーサイドスクリプト
  - データベース初期化 (setupDatabase)
  - セッショントークン発行・検証
  - 排他制御 (LockService)
  - プレイヤー管理、島管理、マルチプレイ同期API
  - 探索、インベントリ、共有倉庫、クラフト、建築、時間進行、イベント、チャット、救助、脱出API
- `index.html` : メインHTMLテンプレート
  - タイトル、プレイヤー登録、島作成、参加、ロビー、HUD、各タブ画面、モーダル、リザルト画面
- `css.html` : ダークファンタジーRPGスタイルシート
  - レスポンシブ対応、ステータスゲージ、アイコン、カード、アニメーション、トースト
- `js.html` : クライアントJavaScript
  - 状態管理、API非同期呼び出し、Web Audio API効果音、ポーリング同期 (3〜5秒)、ローカルストレージ連携
- `README.md` : 初心者向け導入ガイド（スプレッドシート準備、GAS配置、デプロイ、テスト手順）

## 2. データベース設計 (Googleスプレッドシート 12シート)
1. **Players**
   `player_id`, `name`, `role`, `hp`, `hunger`, `thirst`, `stamina`, `status`, `created_at`, `updated_at`, `auth_token`
2. **Islands**
   `island_id`, `island_name`, `invite_code`, `host_player_id`, `day`, `time_period`, `island_status`, `created_at`, `updated_at`
3. **Members**
   `island_id`, `player_id`, `role`, `member_status`, `joined_at`, `last_seen`
4. **Inventory**
   `player_id`, `item_id`, `quantity`
5. **SharedInventory**
   `island_id`, `item_id`, `quantity`
6. **Items**
   `item_id`, `item_name`, `category`, `description`, `effect`
7. **Recipes**
   `recipe_id`, `result_item_id`, `wood`, `stone`, `other_material`, `description`
8. **Events**
   `event_id`, `event_name`, `description`, `effect_type`, `effect_value`
9. **Logs**
   `log_id`, `island_id`, `player_id`, `day`, `action`, `created_at`
10. **Messages**
    `message_id`, `island_id`, `player_id`, `player_name`, `message`, `created_at`
11. **Buildings**
    `building_id`, `island_id`, `building_type`, `level`, `created_by`, `created_at`
12. **Achievements**
    `achievement_id`, `player_id`, `achievement_name`, `unlocked_at`

## 3. 処理の流れとセキュリティ
1. **認証フロー**
   - プレイヤー登録時/ログイン時にUUIDによる `auth_token` を発行し `Players` シートに記録。
   - クライアントは `localStorage` に保持し、全APIリクエスト時に `(playerId, token)` を送信。
   - サーバー側でトークンを厳密に照合。不一致時は操作を拒否。
2. **排他制御 (Concurrency Control)**
   - `LockService.getScriptLock()` により、島作成、参加人数チェック（最大4人）、アイテム移動、クラフト、時間進行等の処理を排他制御。
3. **ゲームループ**
   - クライアントは3〜5秒ごとに `getIslandState` をポーリング。
   - 差分ログ・チャット・メンバー情報・共有倉庫・建築物情報を同期。
   - 各行動（探索、クラフト、建築、時間進行）はサーバー側でバリデーションと計算を行い、結果をDBに永続化。
