// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// HEURÉKA — Configuration globale partagée
// NE JAMAIS modifier SPREADSHEET_ID sans autorisation explicite
// NE JAMAIS lire ou écrire cet ID depuis/vers localStorage
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const HEUREKA_CONFIG = {
  SPREADSHEET_ID:  '1NVUNaS5qOo6PMZvSfh3MLDuJCmsUp5qu10967SaGipw',
  ANTHROPIC_API_KEY: '',
  APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbw7Z9CuFoUrkv5KG_wRgo5Zvm-TXnr2gu66S1kx9y9SBnthkxGMuFpj087TWtIKwuDLhw/exec',
  SYNC_INTERVAL_MS: 5000,
  SESSION_DURATION_H: 8,

  APPS: {
    ADMIN:    'https://daphou8844.github.io/gestions-heureka/admin.html',
    PIPELINE: 'https://daphou8844.github.io/heureka-pipeline/pipeline-heureka.html',
    PUNCH:    'https://daphou8844.github.io/gestions-heureka/punch.html',
    MARKETING: 'https://daphou8844.github.io/heureka-marketing-app/frontend/index.html'
  }
};
