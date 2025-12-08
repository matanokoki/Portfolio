# Business Automation Tool: WBS & Task Management System

## 1. Project Overview
ベンチャーキャピタル（VC）での実務において、煩雑化していたタスク管理を効率化するために開発した自動化システム。 Google Apps Script (GAS) を活用し、進捗の自動更新からSlackへのリマインド通知、複数シートの統合までをワンストップで自動化しました。

**解決した課題:**
* 手動更新による更新漏れと、進捗確認のコミュニケーションコスト（「あれどうなった？」）の削減。
* 複数のプロジェクトシートを横断して確認する手間を解消。

## 2. Key Features
このシステムは以下の3つの主要モジュールで構成されています。

### A. Task Status Automation & Slack Notification (`wbs_automation_logic.js`)
* **Status Auto-Update:**
    * 開始日や締切日を過ぎたタスクのステータス（「未着手」→「実施中」、「実施中」→「遅延」）を自動更新。
* **Daily Digest:**
    * 毎朝9時に、その日開始のタスクや遅延タスクをSlackへダイジェスト通知。
    * 担当者ごとのメンション機能を実装し、見落としを防止。

### B. Sheet Integration System (`wbs_automation_logic.js`)
* **Data Aggregation:**
    * 複数のプロジェクトシート（タブ）に散らばるタスクを、「統合シート」に自動集約。
* **Visualization:**
    * 集約したデータを元に、「誰が」「どのくらい」タスクを抱えているかを可視化するダッシュボード機能。

### C. UX Optimization Utilities (`sheet_list_manager.js`)
* **Auto Indexing:**
    * シートの増減を検知し、目次（シート一覧）を自動生成・更新。
    * ユーザーが目的のプロジェクトシートへ1クリックでアクセスできるUXを実現。

## 3. Technology Stack
* **Language:** Google Apps Script (JavaScript based)
* **Services:** Google Spreadsheet, Slack API (Incoming Webhook)
* **Development:** Gemini (Generative AI) utilizing for logic generation and debugging.

## 4. Usage
(※社外秘情報を含むため、コードの一部はマスキング処理を行っています)
1. `wbs_automation_logic.js`: Time-based triggers (e.g., daily at 9:00 AM) execute the main logic.
2. `sheet_list_manager.js`: Executed when a sheet change event occurs.