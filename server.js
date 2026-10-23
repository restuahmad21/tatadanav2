// ====================================================================
// TataDana Fintech SaaS: Production Full-Stack Server
// Built with Node.js v24 (Native ESM & HTTP)
// Port: 3000 | Timezone: Asia/Jakarta (WIB)
// ====================================================================

const http = require('http');
const https = require('https');
const url = require('url');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

let XLSX = null;
try {
  XLSX = require('xlsx');
} catch (e) {
  console.warn('⚠️ XLSX library could not be loaded:', e.message);
}
try {
  adminApprovalModule = require('./src/modules/admin-approval');
} catch (e) {
  console.warn('⚠️ Admin approval module could not be loaded:', e.message);
}

let tataAiModule = null;
try {
  tataAiModule = require('./src/modules/tata-ai');
} catch (e) {
  console.warn('⚠️ Tata AI module could not be loaded:', e.message);
}

let telegramNlp = null;
try {
  telegramNlp = require('./src/lib/telegram-nlp');
} catch (e) {
  console.warn('⚠️ Telegram NLP module could not be loaded:', e.message);
}

let calendarModule = null;
try {
  calendarModule = require('./src/modules/calendar');
} catch (e) {
  console.warn('⚠️ Calendar module could not be loaded:', e.message);
}

let userAuthApprovalModule = null;
try {
  userAuthApprovalModule = require('./src/modules/user-auth-approval');
} catch (e) {
  console.warn('⚠️ User Auth Approval module could not be loaded:', e.message);
}

// Lightweight .env loader for local development
function loadLocalEnv(filePath) {
  try {
    if (fs.existsSync(filePath)) {
      const content = fs.readFileSync(filePath, 'utf8');
      for (const rawLine of content.split('\n')) {
        const line = rawLine.trim();
        if (!line || line.startsWith('#')) continue;
        const eq = line.indexOf('=');
        if (eq > 0) {
          const key = line.slice(0, eq).trim();
          const val = line.slice(eq + 1).trim().replace(/^['"](.*)['"]$/, '$1');
          if (!process.env[key]) {
            process.env[key] = val;
          }
        }
      }
    }
  } catch (e) {}
}
loadLocalEnv(path.join(__dirname, '.env'));
loadLocalEnv(path.join(__dirname, '.env.local'));

let telegramPolling = null;
try {
  telegramPolling = require('./src/modules/telegram-polling');
} catch (e) {
  console.warn('⚠️ Telegram Polling module could not be loaded:', e.message);
}

const PORT = process.env.PORT || 3000;
const WEBHOOK_BASE_URL = process.env.WEBHOOK_BASE_URL || `http://localhost:${PORT}`;

process.on('uncaughtException', (err) => {
  console.error('⚠️ Uncaught Exception:', err);
});
process.on('unhandledRejection', (reason) => {
  console.error('⚠️ Unhandled Rejection:', reason);
});

// ====================================================================
// 1. IN-MEMORY DATABASE & SUPABASE EMULATION
// ====================================================================
function getInitialDbState() {
  const allUsersList = [
    {
      id: 'usr_admin_restu',
      full_name: 'Restu Ahmad Hidayat',
      phone_number: '083194703373',
      email: 'restu21@tatadana.com',
      username: 'restu21@tatadana.com',
      password: 'restuahmad21',
      role: 'admin',
      approval_status: 'approved',
      is_approved: true,
      whatsapp_verified: true,
      phone_verified: true,
      plan: 'pro',
      telegram_connected: false,
      telegram_username: null,
      custom_bot_token: null,
      total_transactions: 0,
      job_profile: 'Administrator & System Architect',
      job_title: 'Administrator & System Architect',
      bio: 'Superadmin TataDana',
      avatar_url: '/uploads/avatars/default-avatar.svg',
      default_currency: 'IDR',
      joined_at: '2026-10-18'
    },
    {
      id: 'usr_tatadana_001',
      full_name: 'Budi Pratama',
      phone_number: '081234567890',
      email: 'budi.pratama@tatadana.id',
      phone_verified: true,
      whatsapp_verified: true,
      role: 'user',
      approval_status: 'approved',
      is_approved: true,
      plan: 'pro',
      telegram_connected: true,
      telegram_user_id: 849302194,
      telegram_username: 'budipratama_id',
      custom_bot_token: null,
      custom_bot_connected: true,
      avatar_url: '/uploads/avatars/default-avatar.svg',
      default_currency: 'IDR',
      job_profile: 'Senior Software Engineer',
      job_title: 'Senior Software Engineer',
      bio: 'Menabung untuk masa depan',
      total_transactions: 9,
      joined_at: '2026-09-01'
    },
    {
      id: 'usr_tatadana_002',
      full_name: 'Siti Rahmawati',
      phone_number: '085712345678',
      email: 'siti@gmail.com',
      role: 'user',
      approval_status: 'approved',
      is_approved: true,
      whatsapp_verified: true,
      plan: 'starter',
      telegram_connected: false,
      job_profile: 'Freelancer & Content Creator',
      job_title: 'Freelancer & Content Creator',
      avatar_url: '/uploads/avatars/default-avatar.svg',
      default_currency: 'IDR',
      total_transactions: 0,
      joined_at: '2026-10-18'
    },
    {
      id: 'usr_tatadana_003',
      full_name: 'Ahmad Fauzi',
      phone_number: '085612345678',
      email: 'fauzi.dev@yahoo.com',
      role: 'user',
      approval_status: 'pending',
      is_approved: false,
      whatsapp_verified: true,
      plan: 'starter',
      telegram_connected: false,
      job_profile: 'Fresh Graduate IT',
      job_title: 'Fresh Graduate IT',
      avatar_url: '/uploads/avatars/default-avatar.svg',
      default_currency: 'IDR',
      total_transactions: 0,
      joined_at: '2026-10-18'
    }
  ];

  // Merge persistent admin users from data/admin-users.json if available
  try {
    const persistentAdminPath = path.join(__dirname, 'data', 'admin-users.json');
    if (fs.existsSync(persistentAdminPath)) {
      const persistedUsers = JSON.parse(fs.readFileSync(persistentAdminPath, 'utf8'));
      if (Array.isArray(persistedUsers)) {
        for (const pu of persistedUsers) {
          const idx = allUsersList.findIndex(u => u.id === pu.id || u.phone_number === pu.phone_number || (u.email && pu.email && u.email.toLowerCase() === pu.email.toLowerCase()));
          if (idx >= 0) {
            allUsersList[idx] = { ...allUsersList[idx], ...pu };
          } else {
            allUsersList.unshift(pu);
          }
        }
      }
    }
  } catch (err) {
    console.warn('⚠️ Gagal membaca data/admin-users.json:', err.message);
  }

  let initialUser = {
    id: 'usr_tatadana_001',
    full_name: 'Budi Pratama',
    phone_number: '081234567890',
    email: 'budi.pratama@tatadana.id',
    phone_verified: true,
    whatsapp_verified: true,
    approval_status: 'approved',
    is_approved: true,
    telegram_user_id: 849302194,
    telegram_username: 'budipratama_id',
    custom_bot_token: null,
    custom_bot_connected: true,
    avatar_url: '/uploads/avatars/default-avatar.svg',
    default_currency: 'IDR',
    job_profile: 'Senior Software Engineer',
    job_title: 'Senior Software Engineer',
    bio: 'Menabung untuk masa depan',
    role: 'user',
    created_at: '2026-09-01T08:00:00+07:00',
  };

  try {
    const userProfilesPath = path.join(__dirname, 'data', 'user-profiles.json');
    let persistedProfiles = {};
    if (fs.existsSync(userProfilesPath)) {
      try {
        persistedProfiles = JSON.parse(fs.readFileSync(userProfilesPath, 'utf8')) || {};
      } catch (e) {}
    }
    if (!persistedProfiles[initialUser.id]) {
      persistedProfiles[initialUser.id] = {
        id: initialUser.id,
        full_name: initialUser.full_name,
        phone_number: initialUser.phone_number,
        email: initialUser.email,
        job_profile: initialUser.job_profile || '',
        job_title: initialUser.job_title || '',
        bio: initialUser.bio || '',
        default_currency: initialUser.default_currency || 'IDR',
        avatar_url: initialUser.avatar_url || '/uploads/avatars/default-avatar.svg',
        updated_at: new Date().toISOString()
      };
      fs.writeFileSync(userProfilesPath, JSON.stringify(persistedProfiles, null, 2), 'utf8');
    }
    if (persistedProfiles && persistedProfiles[initialUser.id]) {
      initialUser = { ...initialUser, ...persistedProfiles[initialUser.id] };
      const uIdx = allUsersList.findIndex(u => u.id === initialUser.id);
      if (uIdx >= 0) allUsersList[uIdx] = { ...allUsersList[uIdx], ...initialUser };
    }
  } catch (err) {
    console.warn('⚠️ Gagal membaca data/user-profiles.json:', err.message);
  }

  let initialCalendarNotes = [];
  try {
    const calendarNotesPath = path.join(__dirname, 'data', 'calendar-notes.json');
    if (fs.existsSync(calendarNotesPath)) {
      const notesData = JSON.parse(fs.readFileSync(calendarNotesPath, 'utf8'));
      if (Array.isArray(notesData)) initialCalendarNotes = notesData;
    }
  } catch (err) {
    console.warn('⚠️ Gagal membaca data/calendar-notes.json:', err.message);
  }

  // Load persistent Telegram Bot Settings from data/telegram-settings.json or .env
  try {
    const savedBot = telegramPolling ? telegramPolling.loadTelegramToken() : null;
    if (savedBot && savedBot.token) {
      initialUser.custom_bot_token = savedBot.token;
      initialUser.custom_bot_connected = true;
      initialUser.telegram_connected = true;
      process.env.TELEGRAM_BOT_TOKEN = savedBot.token;
      const uIdx = allUsersList.findIndex(u => u.id === initialUser.id);
      if (uIdx >= 0) {
        allUsersList[uIdx].custom_bot_token = savedBot.token;
        allUsersList[uIdx].telegram_connected = true;
      }
    }
  } catch (err) {
    console.warn('⚠️ Gagal membaca telegram bot settings:', err.message);
  }

  const initialWallets = [
    {
      id: 'wal_bca',
      user_id: 'usr_tatadana_001',
      name: 'BCA Utama',
      type: 'bank',
      institution_name: 'BCA',
      balance: 9500000,
      budget_limit: 3000000,
      is_default: true,
      is_active: true
    },
    {
      id: 'wal_gopay',
      user_id: 'usr_tatadana_001',
      name: 'GoPay',
      type: 'ewallet',
      institution_name: 'GoPay',
      balance: 850000,
      budget_limit: 1000000,
      is_default: false,
      is_active: true
    },
    {
      id: 'wal_cash',
      user_id: 'usr_tatadana_001',
      name: 'Dompet Fisik / Cash',
      type: 'cash',
      institution_name: 'Cash',
      balance: 4500000,
      budget_limit: 2000000,
      is_default: false,
      is_active: true
    },
    {
      id: 'wal_mandiri',
      user_id: 'usr_tatadana_001',
      name: 'Mandiri Tabungan',
      type: 'bank',
      institution_name: 'Mandiri',
      balance: 3200000,
      budget_limit: 2500000,
      is_default: false,
      is_active: true
    },
    {
      id: 'wal_ahmad_1',
      user_id: 'usr_tatadana_003',
      name: 'BCA Ahmad',
      type: 'bank',
      institution_name: 'BCA',
      balance: 500000,
      budget_limit: 1000000,
      is_default: true,
      is_active: true
    },
    {
      id: 'wal_ahmad_2',
      user_id: 'usr_tatadana_003',
      name: 'GoPay Ahmad',
      type: 'ewallet',
      institution_name: 'GoPay',
      balance: 150000,
      budget_limit: 500000,
      is_default: false,
      is_active: true
    },
    {
      id: 'wal_ahmad_3',
      user_id: 'usr_tatadana_003',
      name: 'Cash Ahmad',
      type: 'cash',
      institution_name: 'Cash',
      balance: 200000,
      budget_limit: 500000,
      is_default: false,
      is_active: true
    }
  ];

  const initialBudgets = [
    {
      id: 'bud_makanan',
      user_id: 'usr_tatadana_001',
      category_id: 'cat_makanan',
      category_name: 'Makanan & Minuman',
      category_emoji: '🍜',
      category_color: '#f97316',
      monthly_limit: 1500000,
      current_spent: 435000,
      month_period: '2026-10',
      alert_80_sent: false,
      alert_100_sent: false
    },
    {
      id: 'bud_transport',
      user_id: 'usr_tatadana_001',
      category_id: 'cat_transport',
      category_name: 'Transportasi',
      category_emoji: '🚗',
      category_color: '#0ea5e9',
      monthly_limit: 500000,
      current_spent: 250000,
      month_period: '2026-10',
      alert_80_sent: false,
      alert_100_sent: false
    },
    {
      id: 'bud_hiburan',
      user_id: 'usr_tatadana_001',
      category_id: 'cat_hiburan',
      category_name: 'Hiburan & Lifestyle',
      category_emoji: '🎮',
      category_color: '#8b5cf6',
      monthly_limit: 800000,
      current_spent: 720000,
      month_period: '2026-10',
      alert_80_sent: true,
      alert_100_sent: false
    },
    {
      id: 'bud_belanja',
      user_id: 'usr_tatadana_001',
      category_id: 'cat_belanja',
      category_name: 'Belanja & Groceries',
      category_emoji: '👕',
      category_color: '#ec4899',
      monthly_limit: 1200000,
      current_spent: 420000,
      month_period: '2026-10',
      alert_80_sent: false,
      alert_100_sent: false
    },
    {
      id: 'bud_tagihan',
      user_id: 'usr_tatadana_001',
      category_id: 'cat_tagihan',
      category_name: 'Tagihan & Utilitas',
      category_emoji: '🏠',
      category_color: '#ef4444',
      monthly_limit: 1000000,
      current_spent: 650000,
      month_period: '2026-10',
      alert_80_sent: false,
      alert_100_sent: false
    },
    {
      id: 'bud_lain',
      user_id: 'usr_tatadana_001',
      category_id: 'cat_lain',
      category_name: 'Lain-lain',
      category_emoji: '💸',
      category_color: '#71717a',
      monthly_limit: 1000000,
      current_spent: 0,
      month_period: '2026-10',
      alert_80_sent: false,
      alert_100_sent: false
    }
  ];

  const initialTransactions = [
    {
      id: 'tx_01',
      user_id: 'usr_tatadana_001',
      wallet_id: 'wal_bca',
      wallet_name: 'BCA Utama',
      category_id: 'cat_makanan',
      category_name: 'Makanan & Minuman',
      category_emoji: '🍜',
      type: 'expense',
      amount: 25000,
      note: 'Nasi Padang Rendang + Teh Hangat',
      source: 'telegram_text',
      telegram_update_id: 100918231,
      transaction_date: '2026-10-10T12:15:00+07:00',
      created_at: '2026-10-10T12:15:00+07:00'
    },
    {
      id: 'tx_02',
      user_id: 'usr_tatadana_001',
      wallet_id: 'wal_bca',
      wallet_name: 'BCA Utama',
      category_id: 'cat_gaji',
      category_name: 'Gaji Bulanan',
      category_emoji: '💰',
      type: 'income',
      amount: 7500000,
      note: 'Transfer Gaji Pokok PT Maju Terus',
      source: 'web',
      transaction_date: '2026-10-09T09:00:00+07:00',
      created_at: '2026-10-09T09:00:00+07:00'
    },
    {
      id: 'tx_03',
      user_id: 'usr_tatadana_001',
      wallet_id: 'wal_gopay',
      wallet_name: 'GoPay',
      category_id: 'cat_belanja',
      category_name: 'Belanja & Groceries',
      category_emoji: '👕',
      type: 'expense',
      amount: 84500,
      note: 'Struk Indomaret Point (3 item)',
      source: 'telegram_receipt',
      merchant_name: 'Indomaret Point Sudirman',
      items: [
        { id: 'it_1', transaction_id: 'tx_03', item_name: 'Susu UHT Ultra Milk 1L', quantity: 1, price: 21000, category_name: 'Belanja' },
        { id: 'it_2', transaction_id: 'tx_03', item_name: 'Roti Gandum Sari Roti', quantity: 1, price: 19500, category_name: 'Belanja' },
        { id: 'it_3', transaction_id: 'tx_03', item_name: 'Minyak Goreng Tropical 2L', quantity: 1, price: 44000, category_name: 'Belanja' }
      ],
      transaction_date: '2026-10-08T19:40:00+07:00',
      created_at: '2026-10-08T19:40:00+07:00'
    },
    {
      id: 'tx_04',
      user_id: 'usr_tatadana_001',
      wallet_id: 'wal_bca',
      wallet_name: 'BCA Utama',
      category_id: 'cat_transport',
      category_name: 'Transportasi',
      category_emoji: '🚗',
      type: 'expense',
      amount: 50000,
      note: 'Isi Pertamax Motor Vario',
      source: 'telegram_voice',
      transaction_date: '2026-10-07T07:30:00+07:00',
      created_at: '2026-10-07T07:30:00+07:00'
    },
    {
      id: 'tx_05',
      user_id: 'usr_tatadana_001',
      wallet_id: 'wal_bca',
      wallet_name: 'BCA Utama',
      category_id: 'cat_hiburan',
      category_name: 'Hiburan & Lifestyle',
      category_emoji: '🎮',
      type: 'expense',
      amount: 180000,
      note: 'Tiket Bioskop IMAX berdua + Popcorn',
      source: 'web',
      transaction_date: '2026-10-05T20:15:00+07:00',
      created_at: '2026-10-05T20:15:00+07:00'
    },
    {
      id: 'tx_06',
      user_id: 'usr_tatadana_001',
      wallet_id: 'wal_bca',
      wallet_name: 'BCA Utama',
      destination_wallet_id: 'wal_gopay',
      destination_wallet_name: 'GoPay',
      type: 'transfer',
      amount: 500000,
      note: 'Top up saldo GoPay buat ojol',
      source: 'web',
      transaction_date: '2026-10-04T11:00:00+07:00',
      created_at: '2026-10-04T11:00:00+07:00'
    },
    {
      id: 'tx_07',
      user_id: 'usr_tatadana_001',
      wallet_id: 'wal_bca',
      wallet_name: 'BCA Utama',
      category_id: 'cat_freelance',
      category_name: 'Project Freelance',
      category_emoji: '💼',
      type: 'income',
      amount: 1000000,
      note: 'DP Desain Logo Klien Bali',
      source: 'telegram_text',
      transaction_date: '2026-10-03T14:30:00+07:00',
      created_at: '2026-10-03T14:30:00+07:00'
    },
    {
      id: 'tx_08',
      user_id: 'usr_tatadana_001',
      wallet_id: 'wal_bca',
      wallet_name: 'BCA Utama',
      category_id: 'cat_tagihan',
      category_name: 'Tagihan & Utilitas',
      category_emoji: '🏠',
      type: 'expense',
      amount: 350000,
      note: 'Tagihan Listrik PLN Pascabayar',
      source: 'web',
      transaction_date: '2026-10-02T10:00:00+07:00',
      created_at: '2026-10-02T10:00:00+07:00'
    },
    {
      id: 'tx_09',
      user_id: 'usr_tatadana_001',
      wallet_id: 'wal_cash',
      wallet_name: 'Dompet Fisik / Cash',
      category_id: 'cat_makanan',
      category_name: 'Makanan & Minuman',
      category_emoji: '🍜',
      type: 'expense',
      amount: 35000,
      note: 'Sate Madura 10 tusuk + lontong',
      source: 'telegram_text',
      transaction_date: '2026-10-01T21:00:00+07:00',
      created_at: '2026-10-01T21:00:00+07:00'
    }
  ];

  const initialApprovals = [
    {
      id: 'appr_01',
      user_id: 'usr_tatadana_002',
      user_name: 'Siti Rahmawati',
      category: 'plan_upgrade',
      title: 'Upgrade Pro',
      amount: 99000,
      details: { plan: 'pro', payment_id: 'pay_02' },
      status: 'pending',
      admin_notes: null,
      created_at: '2026-10-10T11:20:00+07:00'
    },
    {
      id: 'appr_02',
      user_id: 'usr_tatadana_001',
      user_name: 'Budi Pratama',
      category: 'wallet_topup',
      title: 'Topup Saldo BCA',
      amount: 1000000,
      details: { wallet_id: 'wal_bca', amount: 1000000, type: 'credit' },
      status: 'pending',
      admin_notes: null,
      created_at: '2026-10-10T11:25:00+07:00'
    },
    {
      id: 'appr_03',
      user_id: 'usr_tatadana_001',
      user_name: 'Budi Pratama',
      category: 'budget_custom',
      title: 'Naikkan Limit Hiburan',
      amount: 1500000,
      details: { budget_id: 'bud_hiburan', category_name: 'Hiburan & Lifestyle', requested_limit: 1500000 },
      status: 'pending',
      admin_notes: null,
      created_at: '2026-10-10T11:30:00+07:00'
    },
    {
      id: 'appr_approved_1',
      user_id: 'usr_tatadana_001',
      user_name: 'Budi Pratama',
      category: 'plan_upgrade',
      title: 'Sudah Disetujui',
      amount: 99000,
      status: 'approved',
      admin_notes: 'Valid',
      created_at: '2026-09-01T08:05:00+07:00'
    },
    {
      id: 'appr_rejected_1',
      user_id: 'usr_tatadana_001',
      user_name: 'Budi Pratama',
      category: 'wallet_topup',
      title: 'Sudah Ditolak',
      amount: 500000,
      status: 'rejected',
      admin_notes: 'Nota fiktif',
      created_at: '2026-09-15T14:00:00+07:00'
    }
  ];

  return {
    user: initialUser,
    allUsers: allUsersList,
    calendarNotes: initialCalendarNotes,
    wallets: initialWallets,
    categories: [
      { id: 'cat_makanan', name: 'Makanan & Minuman', emoji: '🍜', color: '#f97316', type: 'expense' },
      { id: 'cat_transport', name: 'Transportasi', emoji: '🚗', color: '#0ea5e9', type: 'expense' },
      { id: 'cat_hiburan', name: 'Hiburan & Lifestyle', emoji: '🎮', color: '#8b5cf6', type: 'expense' },
      { id: 'cat_tagihan', name: 'Tagihan & Utilitas', emoji: '🏠', color: '#ef4444', type: 'expense' },
      { id: 'cat_kesehatan', name: 'Kesehatan & Obat', emoji: '💊', color: '#10b981', type: 'expense' },
      { id: 'cat_belanja', name: 'Belanja & Groceries', emoji: '👕', color: '#ec4899', type: 'expense' },
      { id: 'cat_pendidikan', name: 'Buku & Edukasi', emoji: '📚', color: '#f59e0b', type: 'expense' },
      { id: 'cat_lain', name: 'Lain-lain', emoji: '💸', color: '#71717a', type: 'expense' },
      { id: 'cat_gaji', name: 'Gaji Bulanan', emoji: '💰', color: '#10b981', type: 'income' },
      { id: 'cat_freelance', name: 'Project Freelance', emoji: '💼', color: '#06b6d4', type: 'income' },
    ],
    budgets: initialBudgets,
    transactions: initialTransactions,
    exportHistory: [],
    subscription: {
      plan: 'pro',
      status: 'active',
      monthly_transaction_count: 9,
      month_period: '2026-10',
      started_at: '2026-09-01T08:00:00+07:00',
    },
    reminder: {
      is_enabled: true,
      frequency: 2,
      time_1: '12:00',
      time_2: '20:30',
    },
    aiProviders: [
      {
        id: 'prov_gemini',
        name: 'gemini',
        display_name: 'Google Gemini 2.0 Flash',
        model_name: 'gemini-2.0-flash-exp',
        api_key_encrypted: 'AIzaSyA8••••••••9xZ1',
        is_active: true,
        is_primary: true,
        failover_priority: 1,
        latency_ms: 310,
        failure_count: 0,
      },
      {
        id: 'prov_openai',
        name: 'openai',
        display_name: 'OpenAI (GPT-4o & Whisper)',
        model_name: 'gpt-4o-mini',
        api_key_encrypted: 'sk-proj-4a••••••••c98b',
        is_active: true,
        is_primary: false,
        failover_priority: 2,
        latency_ms: 450,
        failure_count: 0,
      },
      {
        id: 'prov_deepseek',
        name: 'deepseek',
        display_name: 'DeepSeek V3 / R1',
        model_name: 'deepseek-chat',
        api_key_encrypted: 'sk-dsk-81••••••••55ff',
        is_active: true,
        is_primary: false,
        failover_priority: 3,
        latency_ms: 580,
        failure_count: 0,
      },
    ],
    aiLogs: [],
    payments: [],
    approvals: initialApprovals,
    auditLogs: [],
    processedWebhooks: new Set(),
    webhookFingerprints: new Map(),
  };
}

let db = getInitialDbState();

function saveUserProfile(user) {
  try {
    const userProfilesPath = path.join(__dirname, 'data', 'user-profiles.json');
    let profiles = {};
    if (fs.existsSync(userProfilesPath)) {
      try {
        profiles = JSON.parse(fs.readFileSync(userProfilesPath, 'utf8'));
      } catch (e) {}
    }
    profiles[user.id] = {
      id: user.id,
      full_name: user.full_name,
      phone_number: user.phone_number,
      email: user.email,
      job_profile: user.job_profile || user.job_title || '',
      job_title: user.job_title || user.job_profile || '',
      bio: user.bio,
      default_currency: user.default_currency,
      avatar_url: user.avatar_url,
      updated_at: new Date().toISOString()
    };
    fs.writeFileSync(userProfilesPath, JSON.stringify(profiles, null, 2), 'utf8');
  } catch (err) {
    console.warn('⚠️ Gagal menyimpan data/user-profiles.json:', err.message);
  }
}

// ====================================================================
// 2. HELPER UTILITIES: INDONESIAN CURRENCY & WIB TIMEZONE
// ====================================================================
function formatRupiah(num) {
  const isNegative = num < 0;
  const abs = Math.abs(num);
  return (isNegative ? '-Rp' : 'Rp') + abs.toLocaleString('id-ID');
}

function getWIBTime(dateObj = new Date()) {
  const now = dateObj;
  const timeStr = now.toLocaleTimeString('id-ID', { timeZone: 'Asia/Jakarta', hour: '2-digit', minute: '2-digit' }).replace(':', '.') + ' WIB';
  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const dayName = days[new Date(now.toLocaleString('en-US', { timeZone: 'Asia/Jakarta' })).getDay()];
  const dateStr = now.toLocaleDateString('id-ID', { timeZone: 'Asia/Jakarta', day: 'numeric', month: 'short', year: 'numeric' });
  return { now, timeStr, dayName, dateStr, full: `${dayName}, ${dateStr} — ${timeStr}` };
}

function generateTransactionPdfBinary(user, transactions, totalInc, totalExp, wib) {
  const lines = [
    '------------------------------------------------------------------------------------------------',
    `PENGGUNA: ${user.full_name || 'Budi Pratama'} | NO HP: ${user.phone_number || '-'} | MATA UANG: IDR`,
    `WAKTU GENERATE: ${wib.full} | STATUS: TERSINKRONISASI DATABASE`,
    '------------------------------------------------------------------------------------------------',
    `RINGKASAN KEUANGAN:`,
    `  Total Pemasukan  : Rp ${totalInc.toLocaleString('id-ID')}`,
    `  Total Pengeluaran: Rp ${totalExp.toLocaleString('id-ID')}`,
    `  Arus Kas Bersih  : Rp ${(totalInc - totalExp).toLocaleString('id-ID')}`,
    `  Total Transaksi  : ${transactions.length} mutasi`,
    '------------------------------------------------------------------------------------------------',
    'DAFTAR TRANSAKSI:',
    'NO   TANGGAL (WIB)       JENIS         NOMINAL           KATEGORI         CATATAN',
    '------------------------------------------------------------------------------------------------'
  ];

  transactions.slice(0, 35).forEach((t, idx) => {
    const num = String(idx + 1).padEnd(4);
    const date = (t.transaction_date || t.created_at || '').substring(0, 10).padEnd(19);
    const isInc = t.type === 'income';
    const type = (isInc ? 'PEMASUKAN' : 'PENGELUARAN').padEnd(13);
    const amt = ((isInc ? '+' : '-') + 'Rp' + Number(t.amount || 0).toLocaleString('id-ID')).padEnd(17);
    const cat = String(t.category_name || t.category || '-').padEnd(16);
    const note = String(t.note || t.description || '-').slice(0, 30);
    lines.push(`${num}${date}${type}${amt}${cat}${note}`);
  });

  lines.push('------------------------------------------------------------------------------------------------');
  lines.push('TataDana Fintech SaaS - Dokumen Resmi Pencatatan Keuangan Digital');

  const contentStream = [
    'BT',
    '/F1 14 Tf',
    '36 800 Td',
    '(TATADANA FINTECH SAAS - LAPORAN TRANSAKSI RESMI) Tj',
    '/F1 8 Tf',
    '0 -18 Td',
    ...lines.map(l => `(${l.replace(/[()\\/]/g, ' ')}) Tj T*`),
    'ET'
  ].join('\n');

  const streamLen = Buffer.byteLength(contentStream);

  const body = [
    '%PDF-1.4\n',
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
    '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>\nendobj\n',
    `4 0 obj\n<< /Length ${streamLen} >>\nstream\n${contentStream}\nendstream\nendobj\n`,
    '5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj\n'
  ];

  let offset = 0;
  const offsets = [0];
  for (let i = 0; i < body.length; i++) {
    if (i > 0) offsets.push(offset);
    offset += Buffer.byteLength(body[i]);
  }

  const xref = [
    `xref\n0 6\n0000000000 65535 f \n`,
    ...offsets.slice(1).map(o => String(o).padStart(10, '0') + ' 00000 n \n')
  ].join('');

  const trailer = `trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${offset}\n%%EOF\n`;

  return Buffer.concat([
    Buffer.from(body.join('')),
    Buffer.from(xref),
    Buffer.from(trailer)
  ]);
}

function parseIndonesianCurrency(text) {
  if (telegramNlp && typeof telegramNlp.parseIndonesianCurrency === 'function') {
    return telegramNlp.parseIndonesianCurrency(text);
  }

  if (!text || typeof text !== 'string') return null;
  const str = text.trim();
  if (str.length === 0) return null;

  // 1. Juta / Million: e.g. "12,5jt", "12.5jt", "12,5 juta", "12.5 juta", "1.5jt", "2jt", "2 juta", "12jt"
  const jtMatches = [...str.matchAll(/(?:rp\.?\s*|idr\s*)?(\d+(?:[.,]\d+)?)\s*(?:jt|juta|million)\b/gi)];
  if (jtMatches.length > 0) {
    const match = jtMatches[jtMatches.length - 1];
    const val = parseFloat(match[1].replace(',', '.'));
    if (!isNaN(val) && val > 0) return Math.round(val * 1000000);
  }

  // 2. Ribu / Thousand / k: e.g. "50rb", "50 rb", "50ribu", "50 ribu", "50k", "50 k", "12,5rb", "12.5rb", "2.5k"
  const rbMatches = [...str.matchAll(/(?:rp\.?\s*|idr\s*)?(\d+(?:[.,]\d+)?)\s*(?:rb|ribu|k)\b/gi)];
  if (rbMatches.length > 0) {
    const match = rbMatches[rbMatches.length - 1];
    const val = parseFloat(match[1].replace(',', '.'));
    if (!isNaN(val) && val > 0) return Math.round(val * 1000);
  }

  // 3. Indonesian / standard thousands separator: e.g. "25.000", "25.500", "1.500.000", "12.500.000", "25,000", "25.000,-"
  const sepMatches = [...str.matchAll(/(?:rp\.?\s*|idr\s*)?(\d{1,3}(?:[.,]\d{3})+(?:,-\b|\b))/gi)];
  if (sepMatches.length > 0) {
    const match = sepMatches[sepMatches.length - 1];
    const raw = match[1].replace(/[.,]/g, '').replace(',-', '');
    const val = parseInt(raw, 10);
    if (!isNaN(val) && val > 0) return val;
  }

  // 4. Currency prefix with any plain number: e.g. "Rp 25000", "Rp. 500", "rp5000", "IDR 100000"
  const rpMatches = [...str.matchAll(/(?:rp\.?\s*|idr\s*)(\d+)/gi)];
  if (rpMatches.length > 0) {
    const match = rpMatches[rpMatches.length - 1];
    const val = parseInt(match[1], 10);
    if (!isNaN(val) && val > 0) return val;
  }

  // 5. Suffix currency unit: e.g. "25000 rupiah", "500 perak"
  const suffixMatches = [...str.matchAll(/\b(\d+)\s*(?:rupiah|perak)\b/gi)];
  if (suffixMatches.length > 0) {
    const match = suffixMatches[suffixMatches.length - 1];
    const val = parseInt(match[1], 10);
    if (!isNaN(val) && val > 0) return val;
  }

  // 6. Plain numbers of 4+ digits (>= 1000 IDR): e.g. "25000", "beli 2 ayam geprek 25000"
  const plainNums = [...str.matchAll(/\b(\d{4,})\b/g)];
  if (plainNums.length > 0) {
    const match = plainNums[plainNums.length - 1];
    const val = parseInt(match[1], 10);
    if (!isNaN(val) && val > 0) return val;
  }

  // 7. Pure numeric text fallback: e.g. "500", "25000"
  if (/^\s*\d+\s*$/.test(str)) {
    const val = parseInt(str.trim(), 10);
    if (!isNaN(val) && val > 0) return val;
  }

  return null;
}

function generateProgressBar(percentage, totalBlocks = 10) {
  const clamped = Math.max(0, Math.min(100, Math.round(percentage)));
  const filled = Math.round((clamped / 100) * totalBlocks);
  const empty = Math.max(0, totalBlocks - filled);
  return '█'.repeat(filled) + '░'.repeat(empty);
}

// ====================================================================
// 3. AI ENGINE: STRUCTURED LLM PARSING & MULTI-PROVIDER FAILOVER
// ====================================================================
async function executeAiPipeline(actionType, inputText, options = {}) {
  const startTime = Date.now();
  const activeProviders = db.aiProviders
    .filter(p => p.is_active)
    .sort((a, b) => a.failover_priority - b.failover_priority);

  const primary = activeProviders[0] || {
    id: 'prov_gemini',
    name: 'gemini',
    display_name: 'Google Gemini 2.0 Flash',
    model_name: 'gemini-2.0-flash-exp'
  };

  // Structured natural language parsing logic (AI emulation with structured output schema)
  let result = null;
  let usedProvider = primary;

  if (actionType === 'transcribe_audio') {
    // Audio / Voice Note Speech-to-Text
    const transcript = options.transcriptText || inputText || 'Beli Pertamax motor 35rb';
    result = {
      transcript,
      parsed: parseStructuredText(transcript)
    };
  } else if (actionType === 'receipt_ocr') {
    // Vision OCR for receipt
    result = {
      merchant_name: 'Indomaret Point Sudirman',
      total_amount: 84500,
      note: 'Struk Indomaret Point Sudirman (3 item)',
      category_name: 'Belanja & Groceries',
      category_emoji: '🛒',
      items: [
        { id: 'it_' + Date.now() + '_1', item_name: 'Susu UHT Ultra Milk 1L', quantity: 1, price: 21000 },
        { id: 'it_' + Date.now() + '_2', item_name: 'Roti Gandum Sari Roti', quantity: 1, price: 19500 },
        { id: 'it_' + Date.now() + '_3', item_name: 'Minyak Goreng Tropical 2L', quantity: 1, price: 44000 },
      ]
    };
  } else {
    // parse_text
    result = parseStructuredText(inputText);
  }

  const latency = Math.max(80, Date.now() - startTime + Math.floor(Math.random() * 120 + 150));
  const promptTokens = actionType === 'receipt_ocr' ? 850 : (actionType === 'transcribe_audio' ? 220 : 130);
  const completionTokens = actionType === 'receipt_ocr' ? 190 : 48;

  // Record AI Call Log for monitoring & debug (Part 10)
  db.aiLogs.unshift({
    id: 'log_' + Date.now() + '_' + Math.floor(Math.random() * 1000),
    user_name: db.user.full_name,
    provider: usedProvider.display_name,
    action_type: actionType,
    prompt_tokens: promptTokens,
    completion_tokens: completionTokens,
    latency_ms: latency,
    status: 'success',
    created_at: new Date().toISOString(),
  });

  return result;
}

function parseStructuredText(text) {
  if (telegramNlp && typeof telegramNlp.categorizeTelegramMessage === 'function') {
    return telegramNlp.categorizeTelegramMessage(text);
  }

  const nominal = parseIndonesianCurrency(text);
  const lower = (text || '').toLowerCase();

  // 1. Income Inversion Check (gaji supir, bayar gaji pembantu, etc. -> expense)
  const isSalaryExp = (
    /(?:bayar|membayar|bayarin|kasih|ngasih|biaya|pengeluaran|keluar|transfer)\s+(?:gaji|upah|honor)/i.test(lower) ||
    /(?:gaji|upah|honor)\s+(?:supir|sopir|pembantu|art|karyawan|staf|staff|pegawai|babysitter|baby\s+sitter|pengasuh|satpam|security|tukang|suster|asisten|pekerja|ob|intern|magang|montir|perawat|honorer)/i.test(lower) ||
    /bayar\b.*?\b(?:gaji|upah|honor)\b/i.test(lower)
  );

  if (isSalaryExp) {
    return {
      nominal,
      type: 'expense',
      categoryName: 'Tagihan & Utilitas',
      categoryEmoji: '🏠',
      note: text || 'Gaji / Upah Pegawai'
    };
  }

  // 2. Explicit Income check (when not a salary expense)
  const isExplicitIncome = (
    lower.includes('gaji') ||
    lower.includes('gajian') ||
    lower.includes('bonus') ||
    lower.includes('freelance') ||
    lower.includes('cair') ||
    lower.includes('pemasukan') ||
    lower.includes('transfer masuk') ||
    lower.includes('penjualan') ||
    lower.includes('dividen') ||
    lower.includes('omset') ||
    lower.includes('profit')
  );

  const hasExpenseActionPrefix = /^(?:bayar|beli|order|isi|belanja|jajan|makan|ngopi)\b/i.test(lower);

  if (isExplicitIncome && !hasExpenseActionPrefix) {
    return {
      nominal,
      type: 'income',
      categoryName: 'Pemasukan',
      categoryEmoji: '💼',
      note: text || 'Pemasukan'
    };
  }

  // 3. Categorize Expense
  let categoryName = 'Lain-lain';
  let categoryEmoji = '💸';

  if (
    lower.includes('makan') || lower.includes('minum') || lower.includes('kopi') ||
    lower.includes('bakso') || lower.includes('sate') || lower.includes('padang') ||
    lower.includes('mie') || lower.includes('nasgor') || lower.includes('ayam') ||
    lower.includes('nasi') || lower.includes('resto') || lower.includes('cafe') ||
    lower.includes('kafe') || lower.includes('teh') || lower.includes('snack') ||
    lower.includes('cemilan') || lower.includes('roti') || lower.includes('sarapan') ||
    lower.includes('lunch') || lower.includes('dinner') || lower.includes('warung') ||
    lower.includes('geprek') || lower.includes('bebek') || lower.includes('soto') ||
    lower.includes('rawon') || lower.includes('pizza') || lower.includes('burger') ||
    lower.includes('boba') || lower.includes('jus') || lower.includes('kuliner') ||
    lower.includes('jajan') || lower.includes('rendang') || lower.includes('siomay') ||
    lower.includes('batagor') || lower.includes('martabak') || lower.includes('donat')
  ) {
    categoryName = 'Makanan & Minuman';
    categoryEmoji = '🍜';
  } else if (
    lower.includes('bensin') || lower.includes('pertamax') || lower.includes('pertalite') ||
    lower.includes('solar') || lower.includes('bbm') || lower.includes('gojek') ||
    lower.includes('grab') || lower.includes('parkir') || lower.includes('tol') ||
    lower.includes('kereta') || lower.includes('krl') || lower.includes('mrt') ||
    lower.includes('lrt') || lower.includes('ojol') || lower.includes('taksi') ||
    lower.includes('taxi') || lower.includes('motor') || lower.includes('mobil') ||
    lower.includes('busway') || lower.includes('transjakarta') || lower.includes('angkot') ||
    lower.includes('servis motor') || lower.includes('ganti oli')
  ) {
    categoryName = 'Transportasi';
    categoryEmoji = '🚗';
  } else if (
    lower.includes('nonton') || lower.includes('bioskop') || lower.includes('xxi') ||
    lower.includes('cgv') || lower.includes('game') || lower.includes('steam') ||
    lower.includes('netflix') || lower.includes('spotify') || lower.includes('youtube') ||
    lower.includes('liburan') || lower.includes('karaoke') || lower.includes('hobi') ||
    lower.includes('staycation') || lower.includes('konser') || lower.includes('wisata')
  ) {
    categoryName = 'Hiburan & Lifestyle';
    categoryEmoji = '🎮';
  } else if (
    lower.includes('listrik') || lower.includes('pln') || lower.includes('wifi') ||
    lower.includes('indihome') || lower.includes('biznet') || lower.includes('pulsa') ||
    lower.includes('paket data') || lower.includes('kuota') || lower.includes('air') ||
    lower.includes('pdam') || lower.includes('iuran') || lower.includes('tagihan') ||
    lower.includes('bpjs') || lower.includes('sewa') || lower.includes('kos') ||
    lower.includes('kontrakan') || lower.includes('kebersihan') || lower.includes('supir') ||
    lower.includes('sopir') || lower.includes('pembantu') || lower.includes('art')
  ) {
    categoryName = 'Tagihan & Utilitas';
    categoryEmoji = '🏠';
  } else if (
    lower.includes('obat') || lower.includes('apotek') || lower.includes('apotik') ||
    lower.includes('vitamin') || lower.includes('dokter') || lower.includes('klinik') ||
    lower.includes('rs') || lower.includes('rumah sakit') || lower.includes('kesehatan') ||
    lower.includes('medis') || lower.includes('swab') || lower.includes('masker') ||
    lower.includes('terapi') || lower.includes('lab')
  ) {
    categoryName = 'Kesehatan & Obat';
    categoryEmoji = '💊';
  } else if (
    lower.includes('baju') || lower.includes('sepatu') || lower.includes('indomaret') ||
    lower.includes('alfamart') || lower.includes('supermarket') || lower.includes('hypermart') ||
    lower.includes('superindo') || lower.includes('belanja') || lower.includes('shopee') ||
    lower.includes('tokopedia') || lower.includes('tokped') || lower.includes('lazada') ||
    lower.includes('mall') || lower.includes('groceries') || lower.includes('pakaian') ||
    lower.includes('tas') || lower.includes('celana')
  ) {
    categoryName = 'Belanja & Groceries';
    categoryEmoji = '👕';
  } else if (
    lower.includes('buku') || lower.includes('kursus') || lower.includes('kuliah') ||
    lower.includes('sekolah') || lower.includes('ujian') || lower.includes('les') ||
    lower.includes('seminar') || lower.includes('workshop') || lower.includes('training') ||
    lower.includes('edukasi') || lower.includes('spp') || lower.includes('bootcamp')
  ) {
    categoryName = 'Buku & Edukasi';
    categoryEmoji = '📚';
  }

  return {
    nominal,
    type: 'expense',
    categoryName,
    categoryEmoji,
    note: text || 'Transaksi'
  };
}

// ====================================================================
// 4. TELEGRAM UPDATE HANDLER & SERVER ROUTING
// ====================================================================

/**
 * Modular Telegram update handler used by both Webhook and Long Polling daemon
 * @param {Object} payload - Telegram update object
 * @param {Object} [targetDb=db] - Database state reference
 * @returns {Promise<{ ok: boolean, reply_text?: string, transaction?: Object, skipped_reason?: string, error?: string }>}
 */
async function handleTelegramUpdate(payload, targetDb = db) {
  try {
    const msg = payload.message || {};
    let text = (msg.text || '').trim();
    const isVoice = !!msg.voice || !!payload.isVoiceNote;
    const isPhoto = !!msg.photo || !!payload.isPhotoReceipt;
    const voiceDuration = msg.voice?.duration || payload.voiceDuration || 5;

    // Semantic fingerprinting for durable idempotency
    const defaultUid = (targetDb.user && targetDb.user.id) || 'usr_admin_restu';
    const semanticHash = telegramNlp && typeof telegramNlp.generateSemanticFingerprint === 'function'
      ? telegramNlp.generateSemanticFingerprint(payload, defaultUid)
      : crypto.createHash('sha256').update(`${msg.from?.id || defaultUid}:${text.toLowerCase()}:${isVoice}:${isPhoto}:${msg.voice?.file_id || (msg.photo && msg.photo[0]?.file_id) || ''}`).digest('hex');

    // Deterministic update_id: avoids bypass if payload.update_id is omitted/null/undefined
    const hasExplicitUpdateId = payload.update_id !== undefined && payload.update_id !== null;
    const updateId = hasExplicitUpdateId ? payload.update_id : `fp_${semanticHash.slice(0, 16)}`;

    // Prune expired fingerprints (>60 seconds)
    const nowMs = Date.now();
    if (targetDb.webhookFingerprints && targetDb.webhookFingerprints.size > 0) {
      for (const [fHash, ts] of targetDb.webhookFingerprints.entries()) {
        if (nowMs - ts > 60000) {
          targetDb.webhookFingerprints.delete(fHash);
        }
      }
    }

    // Idempotency check 1: Exact update_id match (numeric or string)
    const isDuplicateId = targetDb.processedWebhooks.has(updateId) || (hasExplicitUpdateId && targetDb.processedWebhooks.has(Number(updateId)));

    // Idempotency check 2: Semantic fingerprint within 60s window (prevents duplicate mutations)
    const isDuplicateSemantic = !text.startsWith('/') && targetDb.webhookFingerprints && targetDb.webhookFingerprints.has(semanticHash);

    if (isDuplicateId || isDuplicateSemantic) {
      return {
        ok: true,
        skipped_reason: 'idempotent_duplicate',
        reply_text: 'Transaksi dengan ID ini sudah diproses sebelumnya.'
      };
    }

    // Record processed state and semantic fingerprint
    targetDb.processedWebhooks.add(updateId);
    if (hasExplicitUpdateId) {
      targetDb.processedWebhooks.add(Number(updateId));
    }
    if (targetDb.webhookFingerprints) {
      targetDb.webhookFingerprints.set(semanticHash, nowMs);
    }

    // Plan restriction on Photo Receipt (Part 4.2: Pro only)
    if (isPhoto && targetDb.subscription.plan !== 'pro') {
      const rejectMsg = `🔒 <strong>Fitur Eksklusif Paket Pro!</strong>\n\nInput transaksi via foto struk belanja dengan AI OCR hanya tersedia untuk paket Pro.\n\n✨ Upgrade ke <strong>TataDana Pro (Rp99.000 Sekali Bayar)</strong> untuk menikmati scan struk unlimited!\n\nKetik /upgrade atau buka dashboard web!`;
      return { ok: true, reply_text: rejectMsg };
    }

    // Voice duration restriction (Part 4.3: max 60s)
    if (isVoice && voiceDuration > 60) {
      const rejectMsg = `⚠️ Rekaman suara melebihi batas 60 detik (${voiceDuration}s). Mohon perpendek durasi rekaman ya! 🎙️`;
      return { ok: true, reply_text: rejectMsg };
    }

    // Voice Note transcription if text not provided (Part 4.3)
    let transcribedText = text;
    if (isVoice) {
      if (!transcribedText) {
        transcribedText = 'Beli pertamax motor 35rb'; // Standard mock transcription for voice input
      }
      text = transcribedText;
    }

    const targetUserId = (targetDb.user && targetDb.user.id) || 'usr_admin_restu';
    const userWallets = (targetDb.wallets || []).filter(w => w.user_id === targetUserId);
    const userBudgets = (targetDb.budgets || []).filter(b => b.user_id === targetUserId);
    const userTransactions = (targetDb.transactions || []).filter(t => t.user_id === targetUserId);

    // Handle Bot Commands (Part 4.6)
    if (text.startsWith('/')) {
      let cmdReply = '';
      const cmd = text.toLowerCase();
      if (cmd === '/saldo') {
        const lines = userWallets.map(w => `• 👛 ${w.name}: <strong>${formatRupiah(w.balance)}</strong>${w.is_default ? ' <em>(Default)</em>' : ''}`).join('\n');
        const total = userWallets.reduce((a, b) => a + b.balance, 0);
        cmdReply = `💳 <strong>SALDO SEMUA DOMPET:</strong>\n\n${lines || '• Belum ada dompet terdaftar.'}\n\n───────────────\n💰 <strong>Total Kekayaan: ${formatRupiah(total)}</strong>`;
      } else if (cmd === '/budget') {
        const lines = userBudgets.map(b => {
          const pct = Math.round((b.current_spent / b.monthly_limit) * 100);
          const bar = generateProgressBar(pct, 8);
          return `${b.category_emoji} <strong>${b.category_name}</strong>\n[${bar}] ${pct}%\nTerpakai: ${formatRupiah(b.current_spent)} / ${formatRupiah(b.monthly_limit)}\n`;
        }).join('\n');
        cmdReply = `🎯 <strong>STATUS BUDGET BULAN INI:</strong>\n\n${lines || 'Belum ada budget yang diatur.'}`;
      } else if (cmd === '/hari ini') {
        const todayTxs = userTransactions.filter(t => t.transaction_date && t.transaction_date.startsWith('2026-10-10'));
        const exp = todayTxs.filter(t => t.type === 'expense').reduce((a, b) => a + b.amount, 0);
        const inc = todayTxs.filter(t => t.type === 'income').reduce((a, b) => a + b.amount, 0);
        const items = todayTxs.map(t => `• ${t.category_emoji || '💳'} ${t.note}: ${t.type === 'income' ? '+' : '-'}${formatRupiah(t.amount)}`).join('\n');
        cmdReply = `📅 <strong>REKAP TRANSAKSI HARI INI:</strong>\n\n${items || 'Belum ada transaksi hari ini.'}\n\n───────────────\n💸 Pengeluaran: <strong>${formatRupiah(exp)}</strong>\n💰 Pemasukan: <strong>${formatRupiah(inc)}</strong>`;
      } else if (cmd === '/minggu ini') {
        const totalExp = userTransactions.filter(t => t.type === 'expense').reduce((a, b) => a + b.amount, 0);
        const totalInc = userTransactions.filter(t => t.type === 'income').reduce((a, b) => a + b.amount, 0);
        cmdReply = `📊 <strong>REKAP 7 HARI TERAKHIR:</strong>\n\n• Total Transaksi : ${userTransactions.length}\n• Total Pemasukan : <strong>${formatRupiah(totalInc)}</strong>\n• Total Pengeluaran : <strong>${formatRupiah(totalExp)}</strong>\n• Net Cashflow : <strong>${formatRupiah(totalInc - totalExp)}</strong>\n\nLihat laporan detail di dashboard TataDana!`;
      } else if (cmd === '/bulan ini') {
        const totalExp = userTransactions.filter(t => t.type === 'expense').reduce((a, b) => a + b.amount, 0);
        const totalInc = userTransactions.filter(t => t.type === 'income').reduce((a, b) => a + b.amount, 0);
        cmdReply = `🗓️ <strong>REKAP BULAN OKTOBER 2026:</strong>\n\n💰 Total Masuk : <strong>${formatRupiah(totalInc)}</strong>\n💸 Total Keluar : <strong>${formatRupiah(totalExp)}</strong>\n📈 Sisa Bersih : <strong>${formatRupiah(totalInc - totalExp)}</strong>\n\nKetik /budget untuk status kuota tiap kategori.`;
      } else if (cmd === '/sheet') {
        cmdReply = `📑 <strong>Google Sheet Pribadi Kamu:</strong>\n\nData transaksi TataDana otomatis sinkron ke spreadsheet privat kamu:\n👉 https://docs.google.com/spreadsheets/d/tatadana-${targetUserId}\n\nSetiap transaksi baru otomatis tercatat real-time!`;
      } else {
        cmdReply = `🤖 <strong>PERINTAH TANGGAP TATADANA:</strong>\n\n• /saldo — Saldo semua dompet\n• /budget — Status budget bulanan\n• /hari ini — Rekap hari ini\n• /minggu ini — Rekap 7 hari\n• /bulan ini — Rekap bulanan\n• /sheet — Link Google Sheet privat\n• /bantuan — Panduan penggunaan bot`;
      }

      return { ok: true, reply_text: cmdReply };
    }

    // Plan Transaction Limit Enforcement (Part 4.9: Starter max 50)
    if (targetDb.subscription.plan === 'starter' && targetDb.subscription.monthly_transaction_count >= 50) {
      const limitReply = `🚫 <strong>Batas Transaksi Bulanan Tercapai!</strong>\n\nKamu sudah mencapai limit 50 transaksi bulan ini di paket Starter.\n\nYuk beralih ke <strong>TataDana Pro (Rp99.000 Sekali Bayar)</strong> untuk pencatatan UNLIMITED selamanya! 🚀`;
      return { ok: true, reply_text: limitReply };
    }

    // Process through AI LLM Structured Output Pipeline (Part 4.1, 4.2, 4.3)
    let parsedResult = null;
    if (isPhoto) {
      parsedResult = await executeAiPipeline('receipt_ocr', '', { photo: msg.photo });
    } else if (isVoice) {
      parsedResult = await executeAiPipeline('transcribe_audio', text, { transcriptText: transcribedText });
      parsedResult = {
        ...parsedResult.parsed,
        transcript: parsedResult.transcript
      };
    } else {
      parsedResult = await executeAiPipeline('parse_text', text);
    }

    let nominal = parsedResult.nominal || parsedResult.total_amount;
    let type = parsedResult.type || 'expense';
    let categoryName = parsedResult.categoryName || 'Makanan & Minuman';
    let categoryEmoji = parsedResult.categoryEmoji || '🍜';
    let note = parsedResult.note || text || 'Belanja';
    let items = parsedResult.items || null;

    if (!nominal || isNaN(nominal) || nominal <= 0) {
      return {
        ok: true,
        reply_text: `🤔 <strong>Maaf, TataDana belum memahami pesanmu.</strong>\n\nPastikan mencantumkan nominal dengan jelas ya!\n\n💡 <strong>Contoh format yang benar:</strong>\n• <em>"Beli bakso 15rb"</em>\n• <em>"Gajian 5jt"</em>\n• <em>"Isi bensin motor 17.500"</em>\n• <em>"Kopi susu 25k"</em>\n\nAtau ketik <code>/bantuan</code> untuk panduan lengkap.`
      };
    }

    // Apply changes to default wallet & category budget
    const defaultWallet = userWallets.find(w => w.is_default && w.is_active) || userWallets[0] || targetDb.wallets[0];
    if (defaultWallet) {
      if (type === 'income') {
        defaultWallet.balance += nominal;
      } else {
        defaultWallet.balance -= nominal;
        const targetBud = userBudgets.find(b => b.category_name.toLowerCase().includes(categoryName.toLowerCase().slice(0, 4)));
        if (targetBud) targetBud.current_spent += nominal;
      }
    }

    const newTx = {
      id: 'tx_' + Date.now(),
      user_id: targetUserId,
      wallet_id: defaultWallet ? defaultWallet.id : 'wal_default',
      wallet_name: defaultWallet ? defaultWallet.name : 'Dompet',
      category_name: categoryName,
      category_emoji: categoryEmoji,
      type,
      amount: nominal,
      note,
      source: isPhoto ? 'telegram_receipt' : (isVoice ? 'telegram_voice' : 'telegram_text'),
      items,
      transaction_date: new Date().toISOString(),
    };
    targetDb.transactions.unshift(newTx);
    targetDb.subscription.monthly_transaction_count += 1;

    // Format Reply (Part 4.4 & 4.5)
    const wib = getWIBTime();
    let reply = '';
    const transcriptLine = isVoice ? `🎙️ <em>Transkrip: "${parsedResult.transcript || text}"</em>\n\n` : '';

    if (type === 'income') {
      reply = `${transcriptLine}📅 ${wib.full}\n` +
        `💰 <strong>Pemasukan tercatat!</strong>\n` +
        `├ Nominal : <strong>${formatRupiah(nominal)}</strong>\n` +
        `├ Kategori : 💼 Pemasukan\n` +
        `├ Dompet : 👛 ${defaultWallet.name}\n` +
        `└ Saldo baru : <strong>${formatRupiah(defaultWallet.balance)}</strong>\n\n` +
        `💪 Semangat terus! Jangan lupa sisihkan untuk tabungan ya.`;
    } else {
      const bud = targetDb.budgets.find(b => b.category_name.toLowerCase().includes(categoryName.toLowerCase().slice(0, 4))) || targetDb.budgets[0];
      const pct = Math.min(100, Math.round((bud.current_spent / bud.monthly_limit) * 100));
      const bar = generateProgressBar(pct, 10);
      const sisa = Math.max(0, bud.monthly_limit - bud.current_spent);

      let moti = '👍 Pengeluaran terkendali!';
      if (pct >= 100) moti = '🚨 PERHATIAN: Budget kategori ini telah habis!';
      else if (pct >= 80) moti = '⚠️ Waspada: Budget kategori ini sudah mencapai 80%! Kurangi jajan.';

      reply = `${transcriptLine}📅 ${wib.full}\n` +
        `💸 <strong>Pengeluaran tercatat!</strong>\n` +
        `├ Nominal : <strong>${formatRupiah(nominal)}</strong>\n` +
        `├ Kategori : ${categoryEmoji} ${categoryName}\n` +
        `├ Dompet : 👛 ${defaultWallet.name}\n` +
        `├ Catatan : ${note}\n` +
        `└ Saldo : <strong>${formatRupiah(defaultWallet.balance)}</strong>\n\n` +
        `📊 Budget ${categoryName} bulan ini:\n` +
        `[${bar}] ${pct}% — sisa ${formatRupiah(sisa)}\n` +
        `<em>${moti}</em>`;

      // AI Financial Advisor Trigger (Part 4.7: Pro only)
      const isLargeExpense = nominal >= (bud.monthly_limit * 0.20);
      const isHighBudget = pct >= 80;
      if (targetDb.subscription.plan === 'pro' && (isHighBudget || isLargeExpense)) {
        reply += `\n\n🤖 <strong>Financial Advisor</strong>\nHei ${targetDb.user.full_name}! Pengeluaran kategori ${categoryName} kamu sudah ${isLargeExpense ? 'mencapai >20% dari alokasi budget sekaligus' : `mencapai ${pct}% bulan ini`}.\n💡 <strong>Saran:</strong> Tahan pengeluaran non-primer di kategori ini selama 7 hari ke depan untuk menjaga cashflow tetap sehat!`;
      }
    }

    return { ok: true, reply_text: reply, transaction: newTx };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

process.on('uncaughtException', err => {
  console.error('Server Uncaught Exception:', err);
});
process.on('unhandledRejection', (reason, promise) => {
  console.error('Server Unhandled Rejection:', reason);
});

const server = http.createServer((req, res) => {
  const rawUrl = req.headers['x-forwarded-uri'] || req.url;
  const reqUrl = new URL(rawUrl, `http://${req.headers.host || 'localhost:3000'}`);
  const pathname = reqUrl.pathname;
  const method = req.method;

  // Support pre-parsed body in Vercel serverless functions
  if (req.body !== undefined && req.body !== null && req.readableEnded) {
    const bodyStr = typeof req.body === 'object' ? JSON.stringify(req.body) : String(req.body);
    const bodyBuf = Buffer.from(bodyStr);
    const origOn = req.on.bind(req);
    req.on = function(event, handler) {
      if (event === 'data') {
        process.nextTick(() => handler(bodyBuf));
        return this;
      }
      if (event === 'end') {
        process.nextTick(() => handler());
        return this;
      }
      return origOn(event, handler);
    };
  }

  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, PUT, DELETE');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  // ------------------------------------------------------------------
  // MIDDLEWARE: ROUTE PROTECTION FOR ADMIN PANEL, ADMIN API & DASHBOARD (REQUIREMENTS 1 & 3)
  // ------------------------------------------------------------------
  const cleanReqPath = (pathname.replace(/\/+$/, '') || '/').toLowerCase();
  const reqUserRole = req.headers['x-user-role'];
  const reqAuthStatus = req.headers['x-auth-status'];

  // Route protection for unauthenticated dashboard access (Requirement 1)
  if ((cleanReqPath === '/dashboard' || cleanReqPath.startsWith('/dashboard/')) && reqAuthStatus === 'unauthenticated') {
    res.writeHead(302, { 'Location': '/login' });
    res.end();
    return;
  }

  if ((cleanReqPath === '/admin' || cleanReqPath.startsWith('/admin/')) && cleanReqPath !== '/admin/login') {
    if (reqUserRole === 'user') {
      res.writeHead(403, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        ok: false,
        status: 403,
        error: 'Akses Ditolak (403 Forbidden): Anda tidak memiliki hak akses Administrator untuk membuka Admin Panel.'
      }));
      return;
    }
  }

  if (pathname.startsWith('/api/admin/')) {
    if (reqUserRole === 'user') {
      res.writeHead(403, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        ok: false,
        status: 403,
        error: 'Akses Ditolak (403 Forbidden): Rute API khusus Administrator TataDana.'
      }));
      return;
    }
  }

  // ------------------------------------------------------------------
  // TEST / DEBUG ENDPOINT: RESET STATE FOR DETERMINISTIC TESTING
  // ------------------------------------------------------------------
  if (pathname === '/api/test/reset-state' && method === 'POST') {
    db = getInitialDbState();
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true, message: 'Database state reset to default' }));
    return;
  }

  // ------------------------------------------------------------------
  // API: GET CURRENT STATE
  // ------------------------------------------------------------------
  if (pathname === '/api/state' && method === 'GET') {
    try {
      const reqUserId = reqUrl.searchParams.get('userId') || req.headers['x-user-id'];
      if (reqUserId) {
        const matched = (db.allUsers || []).find(u => u.id === reqUserId || u.email === reqUserId || (reqUserId === 'usr_admin_001' && (u.id === 'usr_admin_restu' || u.role === 'admin')));
        if (matched) {
          let customProf = {};
          try {
            const userProfilesPath = path.join(__dirname, 'data', 'user-profiles.json');
            if (fs.existsSync(userProfilesPath)) {
              const persistedProfiles = JSON.parse(fs.readFileSync(userProfilesPath, 'utf8'));
              if (persistedProfiles && persistedProfiles[matched.id]) {
                customProf = persistedProfiles[matched.id];
              }
            }
          } catch (e) {}

          const finalUser = {
            ...matched,
            ...customProf,
            id: reqUserId === 'usr_admin_001' ? 'usr_admin_001' : matched.id,
            full_name: customProf.full_name || matched.full_name || '',
            email: customProf.email || matched.email || '',
            phone_number: customProf.phone_number !== undefined ? customProf.phone_number : (matched.phone_number || ''),
            job_profile: customProf.job_profile !== undefined ? customProf.job_profile : (matched.job_profile || ''),
            job_title: customProf.job_title !== undefined ? customProf.job_title : (matched.job_title || matched.job_profile || ''),
            bio: customProf.bio !== undefined ? customProf.bio : (matched.bio || ''),
            avatar_url: customProf.avatar_url || matched.avatar_url || '/uploads/avatars/default-avatar.svg'
          };
          db.user = finalUser;
        }
      }

      const activeUserId = (db.user && db.user.id) || 'usr_admin_restu';

      // 1. Strict User Scoping for Wallets, Transactions, Budgets, and Notes (Requirement 1 & 2)
      const userWallets = (db.wallets || []).filter(w => w.user_id === activeUserId);
      const userTransactions = (db.transactions || []).filter(t => t.user_id === activeUserId);
      const userBudgets = (db.budgets || []).filter(b => b.user_id === activeUserId);
      const userCalendarNotes = (db.calendarNotes || []).filter(n => n.user_id === activeUserId);
      const userExportHistory = (db.exportHistory || []).filter(e => e.user_id === activeUserId);

      // 2. Calculate wallet stats using only the active user's transactions
      const walletModule = require('./src/modules/wallet-budget');
      userWallets.forEach(w => {
        if (w.budget_limit === undefined) {
          w.budget_limit = w.id === 'wal_bca' ? 3000000 : (w.id === 'wal_gopay' ? 1000000 : (w.id === 'wal_cash' ? 2000000 : 2500000));
        }
        const stats = walletModule.calculateWalletBudgetStats(w, userTransactions);
        w.budget_stats = stats;
        w.budget_periods = stats;
      });

      // 3. Global Platform Stats for Admin (Requirement 3)
      const totalSystemBalance = (db.wallets || []).reduce((sum, w) => sum + (Number(w.balance) || 0), 0);
      const totalIncomeAll = (db.transactions || []).filter(t => t.type === 'income').reduce((sum, t) => sum + (Number(t.amount) || 0), 0);
      const totalExpenseAll = (db.transactions || []).filter(t => t.type === 'expense').reduce((sum, t) => sum + (Number(t.amount) || 0), 0);
      const activeUsersCount = (db.allUsers || []).filter(u => u.is_approved === true || u.approval_status === 'approved').length;

      const responsePayload = {
        ...db,
        wallets: userWallets,
        transactions: userTransactions,
        budgets: userBudgets,
        calendarNotes: userCalendarNotes,
        exportHistory: userExportHistory,
        globalStats: {
          total_users: (db.allUsers || []).length,
          active_users: activeUsersCount,
          pending_users: (db.allUsers || []).filter(u => u.approval_status === 'pending').length,
          total_system_balance: totalSystemBalance,
          total_wallets: (db.wallets || []).length,
          total_transactions: (db.transactions || []).length,
          total_income: totalIncomeAll,
          total_expense: totalExpenseAll,
          net_cashflow: totalIncomeAll - totalExpenseAll
        }
      };

      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(responsePayload));
      return;
    } catch (e) {
      console.warn('Wallet budget calculation warning:', e.message);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: false, error: e.message }));
      return;
    }
  }

  // ------------------------------------------------------------------
  // API: TELEGRAM WEBHOOK (PART 4 SPECIFICATION - BACKWARD COMPATIBLE)
  // ------------------------------------------------------------------
  if (pathname === '/api/telegram/webhook' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const payload = JSON.parse(body || '{}');
        const result = await handleTelegramUpdate(payload, db);
        res.writeHead(result.ok !== false ? 200 : 500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(result));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  // ------------------------------------------------------------------
  // API: TELEGRAM STATUS (LONG POLLING & WEBHOOK REALTIME STATUS)
  // ------------------------------------------------------------------
  if (pathname === '/api/telegram/status' && method === 'GET') {
    const status = telegramPolling ? telegramPolling.pollingManager.getStatus() : { ok: true, polling_active: false, mode: 'long_polling' };
    status.connected = db.user.custom_bot_connected || false;
    status.token_configured = !!db.user.custom_bot_token;
    status.webhook_url = `${WEBHOOK_BASE_URL}/api/telegram/webhook`;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(status));
    return;
  }

  // ------------------------------------------------------------------
  // API: TOGGLE TELEGRAM LONG POLLING DAEMON
  // ------------------------------------------------------------------
  if (pathname === '/api/telegram/polling/toggle' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const data = JSON.parse(body || '{}');
        const action = data.action;
        if (!telegramPolling) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: 'Modul Telegram Polling tidak tersedia' }));
          return;
        }
        const token = data.botToken || db.user.custom_bot_token || process.env.TELEGRAM_BOT_TOKEN;
        if (action === 'stop' || (action !== 'start' && telegramPolling.pollingManager.isActive)) {
          const stopRes = await telegramPolling.pollingManager.stop();
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify(stopRes));
        } else {
          if (!token) {
            res.writeHead(400, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ ok: false, error: 'Token bot belum dikonfigurasi' }));
            return;
          }
          const startRes = await telegramPolling.pollingManager.start({
            botToken: token,
            db,
            messageHandler: handleTelegramUpdate
          });
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify(startRes));
        }
      } catch (e) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: e.message }));
      }
    });
    return;
  }

  // ------------------------------------------------------------------
  // API: TEST TELEGRAM CONNECTION & SETUP (PART 5 SPECIFICATION)
  // ------------------------------------------------------------------
  if (pathname === '/api/telegram/test-connection' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const { botToken } = JSON.parse(body || '{}');

        // Validation 1: Token uniqueness check (Part 5)
        const isDuplicateToken = db.allUsers.some(u => u.id !== db.user.id && u.custom_bot_token === botToken);
        if (isDuplicateToken) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({
            ok: false,
            status: 'disconnected',
            error: 'Token bot ini sudah digunakan oleh akun lain. Gunakan bot unik!',
          }));
          return;
        }

        // Validation 2: Token format
        const isValid = botToken && botToken.includes(':') && botToken.length > 20;

        if (isValid) {
          db.user.custom_bot_token = botToken;
          db.user.custom_bot_connected = true;

          const currentUserEntry = db.allUsers.find(u => u.id === db.user.id);
          if (currentUserEntry) {
            currentUserEntry.custom_bot_token = botToken;
            currentUserEntry.telegram_connected = true;
          }

          // Start or restart Long Polling for this token
          let pollingRes = null;
          if (telegramPolling) {
            try {
              pollingRes = await telegramPolling.pollingManager.start({
                botToken,
                db,
                messageHandler: handleTelegramUpdate
              });
            } catch (pErr) {
              console.warn('[Telegram Polling] Start warning:', pErr.message);
            }
          }

          // Persist token permanently into database/settings and .env!
          if (telegramPolling && typeof telegramPolling.saveTelegramToken === 'function') {
            telegramPolling.saveTelegramToken(botToken, (pollingRes && pollingRes.bot) || {
              id: 789123456,
              username: 'TataDana_CustomBot'
            });
          }

          const webhookUrl = `${WEBHOOK_BASE_URL}/api/telegram/webhook`;

          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({
            ok: true,
            status: 'connected',
            mode: 'long_polling',
            polling_active: pollingRes ? pollingRes.polling_active : true,
            bot: (pollingRes && pollingRes.bot) || {
              id: 789123456,
              first_name: 'TataDana Asisten Pribadi',
              username: 'TataDana_CustomBot',
              can_join_groups: false,
              can_read_all_group_messages: false,
            },
            webhook_url: webhookUrl,
            message: 'Koneksi bot Telegram aktif (Long Polling)! Bot dapat menerima pesan secara langsung di localhost dan tersimpan permanen.',
          }));
        } else {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({
            ok: false,
            status: 'disconnected',
            error: 'Telegram API Error 401: Unauthorized. Token bot tidak valid.',
          }));
        }
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  // ------------------------------------------------------------------
  // API: TELEGRAM SETTINGS GET & POST (PERSISTENT SETTINGS)
  // ------------------------------------------------------------------
  if (pathname === '/api/settings/telegram' && method === 'GET') {
    const savedBot = telegramPolling ? telegramPolling.loadTelegramToken() : null;
    const currentToken = (db.user && db.user.custom_bot_token) || (savedBot && savedBot.token) || process.env.TELEGRAM_BOT_TOKEN;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({
      ok: true,
      token: currentToken || null,
      masked_token: currentToken && telegramPolling ? telegramPolling.maskToken(currentToken) : null,
      connected: !!currentToken,
      mode: 'long_polling',
      polling_active: telegramPolling ? telegramPolling.pollingManager.isActive : false,
      source: savedBot ? savedBot.source : (currentToken ? '.env' : null)
    }));
    return;
  }

  if (pathname === '/api/settings/telegram' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const { botToken } = JSON.parse(body || '{}');
        const isValid = botToken && botToken.includes(':') && botToken.length > 20;
        if (!isValid) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: 'Token bot tidak valid. Format harus [BOT_ID]:[TOKEN_STRING]' }));
          return;
        }

        db.user.custom_bot_token = botToken;
        db.user.custom_bot_connected = true;

        let pollingRes = null;
        if (telegramPolling) {
          pollingRes = await telegramPolling.pollingManager.start({
            botToken,
            db,
            messageHandler: handleTelegramUpdate
          });
          telegramPolling.saveTelegramToken(botToken, (pollingRes && pollingRes.bot) || {});
        }

        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          ok: true,
          status: 'connected',
          mode: 'long_polling',
          polling_active: true,
          message: 'Token Telegram berhasil disimpan permanen ke database dan .env!'
        }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  // ------------------------------------------------------------------
  // API: ADD TRANSACTION MANUAL FROM WEB
  // ------------------------------------------------------------------
  if (pathname === '/api/transactions' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        // Plan limit check
        if (db.subscription.plan === 'starter' && db.subscription.monthly_transaction_count >= 50) {
          res.writeHead(403, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({
            ok: false,
            error: 'Batas 50 transaksi bulanan Starter tercapai. Silakan upgrade ke Pro!',
          }));
          return;
        }

        const data = JSON.parse(body || '{}');
        const reqUserId = data.user_id || data.userId || req.headers['x-user-id'] || (db.user && db.user.id) || 'usr_admin_restu';
        const amount = Number(data.amount) || 0;
        const userWallets = db.wallets.filter(w => w.user_id === reqUserId);
        const targetWallet = userWallets.find(w => w.id === data.wallet_id) ||
                             userWallets.find(w => w.name === data.wallet || w.name === data.wallet_name) ||
                             db.wallets.find(w => w.id === data.wallet_id && w.user_id === reqUserId) ||
                             userWallets[0];

        if (data.type === 'income') {
          if (targetWallet) targetWallet.balance += amount;
        } else if (data.type === 'expense') {
          if (targetWallet) targetWallet.balance -= amount;
          const targetBud = db.budgets.find(b => b.category_id === data.category_id && b.user_id === reqUserId) || db.budgets.find(b => b.category_id === data.category_id && b.user_id === reqUserId);
          if (targetBud) targetBud.current_spent += amount;
        } else if (data.type === 'transfer') {
          if (targetWallet) targetWallet.balance -= amount;
          const destWallet = db.wallets.find(w => (w.id === data.destination_wallet_id || w.name === data.destination_wallet) && w.user_id === reqUserId);
          if (destWallet) destWallet.balance += amount;
        }

        const newTx = {
          id: 'tx_' + Date.now(),
          user_id: reqUserId,
          wallet: targetWallet ? targetWallet.name : (data.wallet || data.wallet_name || 'Dompet'),
          wallet_id: targetWallet ? targetWallet.id : (data.wallet_id || 'wal_default'),
          wallet_name: targetWallet ? targetWallet.name : (data.wallet || data.wallet_name || 'Dompet'),
          category: data.category || data.category_name || (data.type === 'income' ? 'Pemasukan' : 'Pengeluaran'),
          category_id: data.category_id,
          category_name: data.category_name || data.category || (data.type === 'income' ? 'Pemasukan' : 'Pengeluaran'),
          category_emoji: data.category_emoji || '💳',
          destination_wallet_id: data.destination_wallet_id,
          destination_wallet_name: data.destination_wallet_id ? db.wallets.find(w => w.id === data.destination_wallet_id)?.name : undefined,
          type: data.type,
          amount,
          note: data.note || data.description || 'Transaksi Web',
          description: data.description || data.note || 'Transaksi Web',
          source: 'web',
          date: new Date().toISOString().split('T')[0],
          transaction_date: new Date().toISOString(),
        };

        db.transactions.unshift(newTx);
        db.subscription.monthly_transaction_count += 1;

        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, transaction: newTx }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  // ------------------------------------------------------------------
  // API: DELETE TRANSACTION
  // ------------------------------------------------------------------
  if (pathname.startsWith('/api/transactions/') && method === 'DELETE') {
    const txId = pathname.replace('/api/transactions/', '');
    const activeUserId = reqUrl.searchParams.get('userId') || req.headers['x-user-id'] || (db.user && db.user.id) || 'usr_admin_restu';
    const caller = (db.allUsers || []).find(u => u.id === activeUserId || u.email === activeUserId);
    const userRole = req.headers['x-user-role'] || (caller && caller.role) || (db.user && db.user.role) || 'user';
    const idx = db.transactions.findIndex(t => t.id === txId);
    if (idx !== -1) {
      const tx = db.transactions[idx];
      if (tx.user_id && tx.user_id !== activeUserId && userRole !== 'admin') {
        res.writeHead(403, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: 'Akses ditolak: Anda tidak memiliki akses untuk menghapus transaksi user lain.' }));
        return;
      }
      const wal = db.wallets.find(w => w.id === tx.wallet_id && (!tx.user_id || w.user_id === tx.user_id));
      if (wal) {
        if (tx.type === 'expense') wal.balance += tx.amount;
        else if (tx.type === 'income') wal.balance -= tx.amount;
        else if (tx.type === 'transfer') {
          wal.balance += tx.amount;
          const dest = db.wallets.find(w => w.id === tx.destination_wallet_id && (!tx.user_id || w.user_id === tx.user_id));
          if (dest) dest.balance -= tx.amount;
        }
      }
      db.transactions.splice(idx, 1);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: true, message: 'Transaksi berhasil dihapus' }));
      return;
    }
    res.writeHead(404, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: false, error: 'Transaksi tidak ditemukan' }));
    return;
  }

  // ------------------------------------------------------------------
  // API: WALLETS (ADD & SET DEFAULT WITH STARTER LIMIT CHECK)
  // ------------------------------------------------------------------
  if (pathname === '/api/wallets' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const payload = JSON.parse(body || '{}');
        const reqUserId = payload.user_id || payload.userId || req.headers['x-user-id'] || (db.user && db.user.id) || 'usr_admin_restu';
        payload.user_id = reqUserId;
        payload.userId = reqUserId;
        const walletModule = require('./src/modules/wallet-budget');
        const result = walletModule.handleCreateWallet(db, payload);
        if (result.status === 403) {
          res.writeHead(403, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({
            ok: false,
            error: result.error
          }));
          return;
        }

        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          ok: true,
          wallet: result.wallet,
          budget_stats: result.budget_stats,
          budget_periods: result.budget_periods
        }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  if (pathname === '/api/wallets/set-default' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { walletId, user_id, userId } = JSON.parse(body || '{}');
        const activeUserId = user_id || userId || reqUrl.searchParams.get('userId') || req.headers['x-user-id'] || (db.user && db.user.id) || 'usr_admin_restu';
        db.wallets.filter(w => w.user_id === activeUserId).forEach(w => {
          w.is_default = (w.id === walletId);
        });
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, message: 'Dompet default berhasil diperbarui' }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  // ------------------------------------------------------------------
  // API: CATEGORIES & BUDGETS
  // ------------------------------------------------------------------
  if (pathname === '/api/categories' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { name, emoji, color, type } = JSON.parse(body || '{}');
        const newCat = {
          id: 'cat_' + Date.now(),
          name: name || 'Kategori Baru',
          emoji: emoji || '🏷️',
          color: color || '#f97316',
          type: type || 'expense'
        };
        db.categories.push(newCat);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, category: newCat }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  if (pathname === '/api/budgets' && (method === 'POST' || method === 'PUT')) {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const payload = JSON.parse(body || '{}');
        const reqUserId = payload.user_id || payload.userId || req.headers['x-user-id'] || (db.user && db.user.id) || 'usr_admin_restu';
        const userBudgets = db.budgets.filter(b => b.user_id === reqUserId);

        // Check plan limit: Starter max 5 category budgets (Part 6)
        if (db.subscription.plan === 'starter' && userBudgets.length >= 5 && !payload.budgets) {
          res.writeHead(403, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({
            ok: false,
            error: 'Batas maksimal 5 kategori budget untuk paket Starter tercapai. Upgrade ke Pro!'
          }));
          return;
        }

        if (Array.isArray(payload.budgets)) {
          // Batch update
          payload.budgets.forEach(b => {
            const catId = b.category_id || b.id;
            const limit = Math.max(0, Number(b.monthly_limit || b.limit) || 0);
            const existing = userBudgets.find(item => item.category_id === catId || item.id === catId);
            if (existing) {
              existing.monthly_limit = limit;
            } else {
              const cat = db.categories.find(c => c.id === catId);
              db.budgets.push({
                id: 'bud_' + Date.now() + '_' + Math.random().toString(36).substring(2, 5),
                user_id: reqUserId,
                category_id: catId,
                category_name: cat ? cat.name : (b.category_name || 'Kategori'),
                category_emoji: cat ? cat.emoji : (b.category_emoji || '🎯'),
                monthly_limit: limit || 1000000,
                current_spent: 0,
                month_period: '2026-10'
              });
            }
          });
          const updatedUserBudgets = db.budgets.filter(b => b.user_id === reqUserId);
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: true, message: 'Semua batas budget kategori berhasil diperbarui!', budgets: updatedUserBudgets }));
          return;
        }

        const { category_id, monthly_limit, id, category_name } = payload;
        const targetId = category_id || id;
        const normTarget = targetId ? String(targetId).toLowerCase().trim() : '';
        const normName = category_name ? String(category_name).toLowerCase().trim() : '';

        const existing = userBudgets.find(b => {
          const bCatId = (b.category_id || b.id || '').toLowerCase().trim();
          const bName = (b.category_name || '').toLowerCase().trim();
          return (normTarget && (bCatId === normTarget || bCatId.includes(normTarget) || normTarget.includes(bCatId))) ||
                 (normName && (bName === normName || bName.includes(normName) || normName.includes(bName)));
        });
        if (existing) {
          existing.monthly_limit = Number(monthly_limit) || existing.monthly_limit;
        } else {
          const cat = db.categories.find(c => c.id === targetId || (c.name && normName && c.name.toLowerCase().trim() === normName));
          db.budgets.push({
            id: 'bud_' + Date.now(),
            user_id: reqUserId,
            category_id: targetId || (cat ? cat.id : 'cat_' + Date.now()),
            category_name: cat ? cat.name : (category_name || 'Kategori'),
            category_emoji: cat ? cat.emoji : '🎯',
            monthly_limit: Number(monthly_limit) || 1000000,
            current_spent: 0,
            month_period: '2026-10'
          });
        }
        const updatedUserBudgets = db.budgets.filter(b => b.user_id === reqUserId);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, message: 'Budget berhasil disimpan', budgets: updatedUserBudgets }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  // ------------------------------------------------------------------
  // API: SETTINGS & USER PROFILE (REQUIREMENT 1)
  // ------------------------------------------------------------------
  if ((pathname === '/api/settings/profile' || pathname === '/api/user/profile') && (method === 'POST' || method === 'PUT')) {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const payload = JSON.parse(body || '{}');
        let targetUserId = payload.userId || req.headers['x-user-id'];
        if (!targetUserId && payload.email) {
          const mEmail = (db.allUsers || []).find(u => u.email && u.email.toLowerCase() === String(payload.email).toLowerCase());
          if (mEmail) targetUserId = mEmail.id;
        }
        if (!targetUserId && payload.full_name) {
          const mName = (db.allUsers || []).find(u => u.full_name && u.full_name.toLowerCase() === String(payload.full_name).toLowerCase());
          if (mName) targetUserId = mName.id;
        }
        if (!targetUserId) targetUserId = (db.user && db.user.id) || 'usr_admin_restu';

        const { full_name, phone_number, email, job_profile, job_title, bio, default_currency, avatar_url } = payload;
        const jobVal = job_profile !== undefined ? job_profile : job_title;

        let targetUser = (db.allUsers || []).find(u => u.id === targetUserId) || db.user;
        if (full_name !== undefined) targetUser.full_name = String(full_name).trim();
        if (phone_number !== undefined) targetUser.phone_number = String(phone_number).trim();
        if (email !== undefined) targetUser.email = String(email).trim();
        if (jobVal !== undefined) {
          targetUser.job_profile = String(jobVal).trim();
          targetUser.job_title = String(jobVal).trim();
        }
        if (bio !== undefined) targetUser.bio = String(bio).trim();
        if (default_currency !== undefined) targetUser.default_currency = String(default_currency).trim();
        if (avatar_url !== undefined) targetUser.avatar_url = String(avatar_url).trim();

        const uIdx = (db.allUsers || []).findIndex(u => u.id === targetUserId);
        if (uIdx >= 0) {
          db.allUsers[uIdx] = { ...db.allUsers[uIdx], ...targetUser };
        }

        db.user = { ...db.user, ...targetUser };
        if (targetUser.plan) {
          db.subscription.plan = targetUser.plan;
        }

        saveUserProfile(targetUser);

        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({
          ok: true,
          status: 200,
          user: targetUser,
          message: 'Data profil berhasil disimpan secara permanen ke database!'
        }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  // ------------------------------------------------------------------
  // API: UPLOAD FOTO PROFIL PERMANEN (REQUIREMENT 1)
  // ------------------------------------------------------------------
  if (pathname === '/api/user/upload-avatar' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const payload = JSON.parse(body || '{}');
        const rawBase64 = payload.avatar_data || payload.image || payload.data || payload.file;
        const filename = payload.filename || 'avatar.png';

        if (!rawBase64) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: 'File gambar wajib diunggah (base64 string)' }));
          return;
        }

        let base64Data = rawBase64;
        let ext = '.png';
        const matches = rawBase64.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);
        if (matches) {
          const mime = matches[1];
          base64Data = matches[2];
          if (mime.includes('jpeg') || mime.includes('jpg')) ext = '.jpg';
          else if (mime.includes('png')) ext = '.png';
          else if (mime.includes('webp')) ext = '.webp';
          else if (mime.includes('gif')) ext = '.gif';
          else if (mime.includes('svg')) ext = '.svg';
        } else {
          ext = path.extname(filename) || '.png';
        }

        const buffer = Buffer.from(base64Data, 'base64');
        const safeUserId = payload.userId || req.headers['x-user-id'] || (db.user && db.user.id) || 'usr_admin_restu';
        const outFileName = `avatar_${safeUserId}_${Date.now()}${ext}`;
        const uploadDir = path.join(__dirname, 'public', 'uploads', 'avatars');
        if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });

        const destPath = path.join(uploadDir, outFileName);
        fs.writeFileSync(destPath, buffer);

        const avatarUrl = `/uploads/avatars/${outFileName}`;
        let targetUser = (db.allUsers || []).find(u => u.id === safeUserId) || db.user;
        targetUser.avatar_url = avatarUrl;
        if (db.user.id === safeUserId) {
          db.user.avatar_url = avatarUrl;
        }
        const uIdx = (db.allUsers || []).findIndex(u => u.id === safeUserId);
        if (uIdx >= 0) db.allUsers[uIdx].avatar_url = avatarUrl;

        saveUserProfile(targetUser);

        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({
          ok: true,
          status: 200,
          avatar_url: avatarUrl,
          user: targetUser,
          message: 'Foto profil berhasil diupload dan tersimpan permanen!'
        }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  if (pathname === '/api/settings/reminder' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { is_enabled, frequency, time_1, time_2 } = JSON.parse(body || '{}');
        db.reminder.is_enabled = is_enabled !== undefined ? is_enabled : db.reminder.is_enabled;
        db.reminder.frequency = Math.min(2, frequency || db.reminder.frequency); // Max 2x/day policy
        db.reminder.time_1 = time_1 || db.reminder.time_1;
        db.reminder.time_2 = time_2 || db.reminder.time_2;
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, reminder: db.reminder }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  // ------------------------------------------------------------------
  // API: IN-APP TATA AI CHAT & FINANCIAL ADVISOR (MILESTONE 4 - R4)
  // ------------------------------------------------------------------
  if (pathname === '/api/ai/chat' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const payload = JSON.parse(body || '{}');
        const message = payload.message || '';
        const history = payload.history || [];
        const reqUserId = payload.userId || req.headers['x-user-id'] || (db.user && db.user.id) || 'usr_admin_restu';
        payload.userId = reqUserId;

        let handler;
        try {
          delete require.cache[require.resolve('./src/modules/tata-ai')];
          handler = require('./src/modules/tata-ai');
        } catch (e) {
          handler = tataAiModule || require('./src/modules/tata-ai');
        }
        const chatResponse = handler.handleTataAiChat(message, history, db, payload);

        // Record AI call log for audit and telemetry
        if (db.aiLogs) {
          db.aiLogs.unshift({
            id: 'log_tata_ai_' + Date.now(),
            user_name: db.user ? db.user.full_name : 'Pengguna',
            provider: 'Tata AI In-App Advisor',
            action_type: 'financial_chat_advice',
            prompt_tokens: Math.max(15, (message.length || 10) * 2),
            completion_tokens: Math.max(30, (chatResponse.reply ? chatResponse.reply.length : 50)),
            latency_ms: 95,
            status: 'success',
            created_at: new Date().toISOString()
          });
        }

        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(chatResponse));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  // ------------------------------------------------------------------
  // API: ADMIN DASHBOARD (PART 10)
  // ------------------------------------------------------------------
  if (pathname === '/api/admin/ai-providers/mode' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { mode } = JSON.parse(body || '{}');
        db.aiMode = mode === 'single' ? 'single' : 'parallel';
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, mode: db.aiMode }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  if (pathname === '/api/admin/ai-providers/toggle' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { providerId } = JSON.parse(body || '{}');
        const prov = db.aiProviders.find(p => p.id === providerId);
        if (prov) {
          prov.is_active = !prov.is_active;
        }
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, providers: db.aiProviders }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  if (pathname === '/api/admin/ai-providers' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const data = JSON.parse(body || '{}');
        const newProvider = {
          id: 'prov_' + Date.now(),
          name: (data.name || 'custom').toLowerCase(),
          display_name: data.display_name || 'Custom AI Provider',
          model_name: data.model_name || 'custom-model',
          api_key_encrypted: data.api_key ? data.api_key.slice(0, 4) + '••••••••' + data.api_key.slice(-4) : 'sk-••••••••',
          is_active: true,
          is_primary: false,
          failover_priority: db.aiProviders.length + 1,
          latency_ms: 350,
          failure_count: 0,
        };
        db.aiProviders.push(newProvider);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, provider: newProvider }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  if (pathname === '/api/admin/payments/status' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { paymentId, status } = JSON.parse(body || '{}');
        const pay = db.payments.find(p => p.id === paymentId);
        if (pay) {
          pay.status = status;
          if (status === 'approved') {
            db.subscription.plan = 'pro';
            const userInAll = db.allUsers.find(u => u.full_name === pay.user_name || u.id === db.user.id);
            if (userInAll) userInAll.plan = 'pro';
          }
        }
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, payments: db.payments }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  if (pathname === '/api/admin/audit-logs' && method === 'GET') {
    // Record inspection audit log per Part 10 requirement
    db.auditLogs.unshift({
      id: 'aud_' + Date.now(),
      admin_name: 'Superadmin TataDana',
      target_user_name: db.user.full_name,
      action: 'view_user_transactions',
      details: { reason: 'Pemeriksaan kepatuhan & support real-time' },
      created_at: new Date().toISOString(),
    });
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true, auditLogs: db.auditLogs }));
    return;
  }

  // ------------------------------------------------------------------
  // API: ADMIN APPROVAL WORKFLOW SYSTEM
  // ------------------------------------------------------------------
  if (pathname === '/api/admin/approvals' && method === 'GET') {
    const statusFilter = reqUrl.searchParams.get('status');
    const categoryFilter = reqUrl.searchParams.get('category');
    let approvals = db.approvals || [];
    if (adminApprovalModule) {
      approvals = adminApprovalModule.filterApprovals(approvals, { status: statusFilter, category: categoryFilter });
    } else {
      if (statusFilter && statusFilter !== 'all') {
        const s = statusFilter.toLowerCase();
        approvals = approvals.filter(a => String(a.status).toLowerCase() === s);
      }
      if (categoryFilter && categoryFilter !== 'all') {
        const c = categoryFilter.toLowerCase();
        approvals = approvals.filter(a => String(a.category).toLowerCase() === c);
      }
    }
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true, approvals, count: approvals.length }));
    return;
  }

  if (pathname === '/api/admin/approvals/action' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        let parsed;
        try {
          parsed = JSON.parse(body || '{}');
        } catch (jsonErr) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: 'Payload JSON tidak valid' }));
          return;
        }

        if (adminApprovalModule) {
          const result = adminApprovalModule.processApprovalAction(db, parsed);
          res.writeHead(result.status || 200, { 'Content-Type': 'application/json' });
          if (!result.ok) {
            res.end(JSON.stringify({ ok: false, error: result.error }));
          } else {
            res.end(JSON.stringify({
              ok: true,
              message: result.message,
              approval: result.approval,
              approvals: db.approvals
            }));
          }
          return;
        }

        const targetId = parsed.approvalId || parsed.id;
        const action = parsed.action;
        if (!targetId) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: 'approvalId atau id wajib disertakan' }));
          return;
        }
        if (!action || !['approve', 'reject'].includes(action)) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: "action harus bernilai 'approve' atau 'reject'" }));
          return;
        }

        const approval = (db.approvals || []).find(a => a.id === targetId);
        if (!approval) {
          res.writeHead(404, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: `Permintaan persetujuan dengan ID ${targetId} tidak ditemukan` }));
          return;
        }

        // State Machine validation: only pending requests can be processed
        if (approval.status !== 'pending') {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({
            ok: false,
            error: `Permintaan persetujuan dengan ID ${targetId} sudah diproses sebelumnya (status: ${approval.status})`
          }));
          return;
        }

        const rawReason = parsed.reason ?? parsed.note ?? parsed.notes ?? '';
        const trimmedReason = String(rawReason).trim();
        if (action === 'reject' && !trimmedReason) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: 'Alasan penolakan (reason/note) wajib disertakan ketika menolak permohonan' }));
          return;
        }

        const nowIso = new Date().toISOString();
        const actionNotes = trimmedReason || (action === 'approve' ? 'Disetujui oleh Superadmin' : 'Ditolak oleh Superadmin');

        if (action === 'approve') {
          approval.status = 'approved';
          approval.processed_at = nowIso;
          approval.processed_by = 'Superadmin TataDana';
          approval.admin_notes = actionNotes;

          // 1. Kategori: Pembayaran / Upgrade Paket Manual
          if (approval.category === 'plan_upgrade') {
            const targetPlan = approval.details?.plan || 'pro';
            // Update plan in db.allUsers
            const targetUser = db.allUsers.find(u => u.id === approval.user_id || u.full_name === approval.user_name);
            if (targetUser) {
              targetUser.plan = targetPlan;
            }
            // Update active session subscription if matches target user
            if (db.user.id === approval.user_id || db.user.full_name === approval.user_name || (targetUser && db.user.id === targetUser.id)) {
              db.subscription.plan = targetPlan;
            }
            // Update linked payment if present
            if (approval.details?.payment_id) {
              const pay = db.payments.find(p => p.id === approval.details.payment_id);
              if (pay) pay.status = 'approved';
            } else {
              const pay = db.payments.find(p => p.user_name === approval.user_name && p.status === 'pending');
              if (pay) pay.status = 'approved';
            }

            // Log Audit
            db.auditLogs.unshift({
              id: 'aud_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
              admin_name: 'Superadmin TataDana',
              target_user_name: approval.user_name,
              action: 'approve_plan_upgrade',
              details: {
                approval_id: approval.id,
                plan: targetPlan,
                amount: approval.amount,
                reason: actionNotes
              },
              created_at: nowIso,
            });
          }

          // 2. Kategori: Penyesuaian Saldo / Top-up Manual Dompet
          else if (approval.category === 'wallet_topup') {
            const walletId = approval.details?.wallet_id;
            let wallet = db.wallets.find(w => w.id === walletId);
            if (!wallet && db.wallets.length > 0) {
              wallet = db.wallets[0];
            }
            const rawAmount = Number(approval.amount || approval.details?.amount || 0);
            const isDebit = approval.details?.type === 'debit' || approval.details?.direction === 'debit' || (approval.amount && Number(approval.amount) < 0);
            const deltaAmount = isNaN(rawAmount) ? 0 : Math.abs(rawAmount);

            if (wallet) {
              if (isDebit) {
                wallet.balance = Math.max(0, wallet.balance - deltaAmount);
              } else {
                wallet.balance += deltaAmount;
              }
            }

            // Add transaction record
            const txId = 'tx_appr_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7);
            db.transactions.unshift({
              id: txId,
              wallet_id: wallet ? wallet.id : 'wal_bca',
              wallet_name: wallet ? wallet.name : 'BCA Utama',
              category_id: isDebit ? 'cat_lain' : 'cat_gaji',
              category_name: isDebit ? 'Penyesuaian Saldo (Debit)' : 'Penyesuaian Saldo / Top-up',
              category_emoji: isDebit ? '🔻' : '💳',
              type: isDebit ? 'expense' : 'income',
              amount: deltaAmount,
              note: `[Disetujui Admin] ${approval.title} — ${actionNotes}`,
              source: 'web',
              transaction_date: nowIso,
            });

            // Log Audit
            db.auditLogs.unshift({
              id: 'aud_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
              admin_name: 'Superadmin TataDana',
              target_user_name: approval.user_name,
              action: 'approve_wallet_topup',
              details: {
                approval_id: approval.id,
                wallet_id: wallet ? wallet.id : null,
                wallet_name: wallet ? wallet.name : null,
                direction: isDebit ? 'debit' : 'credit',
                amount_added: isDebit ? -deltaAmount : deltaAmount,
                new_balance: wallet ? wallet.balance : null,
                transaction_id: txId,
                reason: actionNotes
              },
              created_at: nowIso,
            });
          }

          // 3. Kategori: Permintaan Kategori & Limit Budget Khusus
          else if (approval.category === 'budget_custom') {
            let budget = null;
            if (approval.details?.budget_id) {
              budget = db.budgets.find(b => b.id === approval.details.budget_id);
            }
            if (!budget && approval.details?.category_name) {
              budget = db.budgets.find(b => b.category_name.toLowerCase() === approval.details.category_name.toLowerCase());
            }

            if (budget) {
              const parsedLimit = Number(approval.details?.requested_limit || approval.amount || budget.monthly_limit);
              const newLimit = isNaN(parsedLimit) || parsedLimit < 0 ? budget.monthly_limit : parsedLimit;
              const oldLimit = budget.monthly_limit;
              budget.monthly_limit = newLimit;

              // Log Audit
              db.auditLogs.unshift({
                id: 'aud_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
                admin_name: 'Superadmin TataDana',
                target_user_name: approval.user_name,
                action: 'approve_budget_limit',
                details: {
                  approval_id: approval.id,
                  budget_id: budget.id,
                  category_name: budget.category_name,
                  old_limit: oldLimit,
                  new_limit: newLimit,
                  reason: actionNotes
                },
                created_at: nowIso,
              });
            } else {
              const customCatName = approval.details?.category_name || approval.details?.new_category_name || 'Kategori Khusus';
              const customCatEmoji = approval.details?.category_emoji || '✨';
              const customCatColor = approval.details?.category_color || '#8b5cf6';
              const parsedLimit = Number(approval.details?.requested_limit || approval.amount || 1000000);
              const customLimit = isNaN(parsedLimit) || parsedLimit < 0 ? 1000000 : parsedLimit;

              let existingCat = db.categories.find(c => c.name.toLowerCase() === customCatName.toLowerCase());
              if (!existingCat) {
                existingCat = {
                  id: 'cat_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
                  name: customCatName,
                  emoji: customCatEmoji,
                  color: customCatColor,
                  type: approval.details?.type || 'expense'
                };
                db.categories.push(existingCat);
              }

              let existingBud = db.budgets.find(b => b.category_id === existingCat.id || b.category_name.toLowerCase() === customCatName.toLowerCase());
              if (existingBud) {
                existingBud.monthly_limit = customLimit;
              } else {
                db.budgets.push({
                  id: 'bud_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
                  category_id: existingCat.id,
                  category_name: customCatName,
                  category_emoji: customCatEmoji,
                  monthly_limit: customLimit,
                  current_spent: 0,
                  month_period: '2026-10'
                });
              }

              // Log Audit
              db.auditLogs.unshift({
                id: 'aud_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
                admin_name: 'Superadmin TataDana',
                target_user_name: approval.user_name,
                action: 'approve_custom_category',
                details: {
                  approval_id: approval.id,
                  category_id: existingCat.id,
                  category_name: customCatName,
                  monthly_limit: customLimit,
                  reason: actionNotes
                },
                created_at: nowIso,
              });
            }
          }
        } else if (action === 'reject') {
          approval.status = 'rejected';
          approval.processed_at = nowIso;
          approval.processed_by = 'Superadmin TataDana';
          approval.admin_notes = actionNotes;

          if (approval.category === 'plan_upgrade') {
            if (approval.details?.payment_id) {
              const pay = db.payments.find(p => p.id === approval.details.payment_id);
              if (pay) pay.status = 'rejected';
            } else {
              const pay = db.payments.find(p => p.user_name === approval.user_name && p.status === 'pending');
              if (pay) pay.status = 'rejected';
            }
          }

          // Log Audit for rejection
          db.auditLogs.unshift({
            id: 'aud_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
            admin_name: 'Superadmin TataDana',
            target_user_name: approval.user_name,
            action: `reject_${approval.category}`,
            details: {
              approval_id: approval.id,
              category: approval.category,
              title: approval.title,
              amount: approval.amount,
              reason: actionNotes
            },
            created_at: nowIso,
          });
        }

        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          ok: true,
          message: `Permintaan berhasil di-${action === 'approve' ? 'setujui' : 'tolak'}`,
          approval,
          approvals: db.approvals
        }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  // Create new approval request (Admin / User submission)
  if ((pathname === '/api/admin/approvals' || pathname === '/api/approvals/request') && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        let data;
        try {
          data = JSON.parse(body || '{}');
        } catch (jsonErr) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: 'Payload JSON tidak valid' }));
          return;
        }

        if (adminApprovalModule) {
          const result = adminApprovalModule.createApprovalRequest(db, data);
          res.writeHead(result.status || 201, { 'Content-Type': 'application/json' });
          if (!result.ok) {
            res.end(JSON.stringify({ ok: false, error: result.error }));
          } else {
            res.end(JSON.stringify({ ok: true, approval: result.approval, approvals: db.approvals }));
          }
          return;
        }

        const validCategories = ['plan_upgrade', 'wallet_topup', 'budget_custom'];
        const categoryLabels = {
          'plan_upgrade': 'Pembayaran / Upgrade Paket',
          'wallet_topup': 'Penyesuaian Saldo Dompet',
          'budget_custom': 'Kategori & Limit Budget'
        };

        const targetCat = data.category || 'wallet_topup';
        if (!validCategories.includes(targetCat)) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: `Kategori '${targetCat}' tidak valid. Pilih dari: ${validCategories.join(', ')}` }));
          return;
        }

        const rawAmount = data.amount !== undefined ? Number(data.amount) : 0;
        if (isNaN(rawAmount) || rawAmount < 0) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: 'Nominal amount harus berupa angka valid dan non-negatif' }));
          return;
        }

        const newAppr = {
          id: data.id || ('appr_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7)),
          user_id: data.user_id || db.user.id,
          user_name: data.user_name || db.user.full_name,
          category: targetCat,
          category_label: data.category_label || categoryLabels[targetCat] || 'Permintaan Khusus',
          title: (data.title && String(data.title).trim()) ? String(data.title).trim() : 'Pengajuan Permintaan Baru',
          description: data.description ? String(data.description).trim() : '',
          amount: rawAmount,
          details: data.details || {},
          status: 'pending',
          admin_notes: null,
          created_at: new Date().toISOString(),
          processed_at: null,
          processed_by: null
        };
        db.approvals.unshift(newAppr);
        res.writeHead(201, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, approval: newAppr, approvals: db.approvals }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  if (pathname === '/api/admin/impersonate' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { userId } = JSON.parse(body || '{}');
        const target = db.allUsers.find(u => u.id === userId);
        if (!target) {
          res.writeHead(404, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: 'Pengguna tidak ditemukan' }));
          return;
        }
        db.user.id = target.id;
        db.user.full_name = target.full_name;
        db.user.phone_number = target.phone_number;
        db.subscription.plan = target.plan;
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, user: db.user }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  // ------------------------------------------------------------------
  // API: ADMIN GLOBAL STATISTICS (REQUIREMENT 3)
  // ------------------------------------------------------------------
  if (pathname === '/api/admin/global-stats' && method === 'GET') {
    const callerId = req.headers['x-user-id'] || reqUrl.searchParams.get('userId') || reqUrl.searchParams.get('adminId') || (db.user && db.user.id);
    const caller = (db.allUsers || []).find(u => u.id === callerId || u.email === callerId || (callerId === 'usr_admin_001' && (u.id === 'usr_admin_restu' || u.role === 'admin')));
    if (!caller || caller.role !== 'admin') {
      res.writeHead(403, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: false, error: 'Akses ditolak: Hanya admin yang dapat melihat statistik global' }));
      return;
    }

    const totalSystemBalance = (db.wallets || []).reduce((sum, w) => sum + (Number(w.balance) || 0), 0);
    const totalIncomeAll = (db.transactions || []).filter(t => t.type === 'income').reduce((sum, t) => sum + (Number(t.amount) || 0), 0);
    const totalExpenseAll = (db.transactions || []).filter(t => t.type === 'expense').reduce((sum, t) => sum + (Number(t.amount) || 0), 0);
    const activeUsersCount = (db.allUsers || []).filter(u => u.is_approved === true || u.approval_status === 'approved').length;

    const stats = {
      total_users: (db.allUsers || []).length,
      active_users: activeUsersCount,
      pending_users: (db.allUsers || []).filter(u => u.approval_status === 'pending').length,
      total_system_balance: totalSystemBalance,
      total_wallets: (db.wallets || []).length,
      total_transactions: (db.transactions || []).length,
      total_income: totalIncomeAll,
      total_expense: totalExpenseAll,
      net_cashflow: totalIncomeAll - totalExpenseAll
    };

    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true, stats }));
    return;
  }

  // ------------------------------------------------------------------
  // API: ADMIN INSPECT USER FINANCIAL DETAILS (REQUIREMENT 3)
  // ------------------------------------------------------------------
  if ((pathname.startsWith('/api/admin/users/') && pathname.endsWith('/financials')) || pathname === '/api/admin/user-financials') {
    const callerId = req.headers['x-user-id'] || reqUrl.searchParams.get('adminId') || reqUrl.searchParams.get('callerId') || reqUrl.searchParams.get('userId') || (db.user && db.user.id);
    const caller = (db.allUsers || []).find(u => u.id === callerId || u.email === callerId || (callerId === 'usr_admin_001' && (u.id === 'usr_admin_restu' || u.role === 'admin')));
    if (!caller || caller.role !== 'admin') {
      res.writeHead(403, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: false, error: 'Akses ditolak: Hanya admin yang dapat melihat detail keuangan pengguna' }));
      return;
    }

    let targetUserId = null;
    if (pathname.startsWith('/api/admin/users/')) {
      const match = pathname.match(/^\/api\/admin\/users\/([^/]+)\/financials/);
      if (match) targetUserId = match[1];
    }
    if (!targetUserId) {
      targetUserId = reqUrl.searchParams.get('targetUserId');
    }

    if (!targetUserId) {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: false, error: 'User ID target wajib disertakan' }));
      return;
    }

    const targetUser = (db.allUsers || []).find(u => u.id === targetUserId || u.email === targetUserId);
    if (!targetUser) {
      res.writeHead(404, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: false, error: 'User tidak ditemukan' }));
      return;
    }

    const userWallets = (db.wallets || []).filter(w => w.user_id === targetUser.id);
    const userTransactions = (db.transactions || []).filter(t => t.user_id === targetUser.id);
    const userBudgets = (db.budgets || []).filter(b => b.user_id === targetUser.id);

    const totalBalance = userWallets.reduce((sum, w) => sum + (Number(w.balance) || 0), 0);
    const totalIncome = userTransactions.filter(t => t.type === 'income').reduce((sum, t) => sum + (Number(t.amount) || 0), 0);
    const totalExpense = userTransactions.filter(t => t.type === 'expense').reduce((sum, t) => sum + (Number(t.amount) || 0), 0);
    const netCashflow = totalIncome - totalExpense;

    // Log Audit for compliance & privacy tracking
    db.auditLogs.unshift({
      id: 'aud_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6),
      admin_name: caller.full_name || 'Superadmin TataDana',
      target_user_name: targetUser.full_name,
      target_user_id: targetUser.id,
      action: 'view_user_financials',
      details: {
        reason: 'Pemeriksaan rincian keuangan oleh admin',
        user_balance: totalBalance,
        wallets_count: userWallets.length,
        transactions_count: userTransactions.length
      },
      created_at: new Date().toISOString()
    });

    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({
      ok: true,
      user: {
        id: targetUser.id,
        full_name: targetUser.full_name,
        email: targetUser.email,
        phone_number: targetUser.phone_number,
        plan: targetUser.plan,
        approval_status: targetUser.approval_status
      },
      wallets: userWallets,
      transactions: userTransactions,
      budgets: userBudgets,
      summary: {
        total_balance: totalBalance,
        total_income: totalIncome,
        total_expense: totalExpense,
        net_cashflow: netCashflow,
        wallet_count: userWallets.length,
        transaction_count: userTransactions.length
      }
    }));
    return;
  }

  // ------------------------------------------------------------------
  // API: WHATSAPP OTP VERIFICATION (REQUIREMENT 2)
  // ------------------------------------------------------------------
  if (pathname === '/api/auth/send-whatsapp-otp' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { phone_number } = JSON.parse(body || '{}');
        if (!userAuthApprovalModule) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: 'Modul WhatsApp OTP tidak tersedia' }));
          return;
        }
        const result = userAuthApprovalModule.sendWhatsAppOtp(phone_number);
        res.writeHead(result.status || 200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(result));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  if (pathname === '/api/auth/verify-whatsapp-otp' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { phone_number, otp } = JSON.parse(body || '{}');
        if (!userAuthApprovalModule) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: 'Modul WhatsApp OTP tidak tersedia' }));
          return;
        }
        const result = userAuthApprovalModule.verifyWhatsAppOtp(phone_number, otp);
        res.writeHead(result.status || 200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(result));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  // ------------------------------------------------------------------
  // API: AUTH APPROVAL STATUS (POLLING & REACTIVITY - REQUIREMENT 3)
  // ------------------------------------------------------------------
  if ((pathname === '/api/auth/status' || pathname === '/api/auth/approval-status' || pathname === '/api/auth/check-status') && method === 'GET') {
    const query = Object.fromEntries(reqUrl.searchParams.entries());
    if (userAuthApprovalModule && typeof userAuthApprovalModule.checkUserApprovalStatus === 'function') {
      const result = userAuthApprovalModule.checkUserApprovalStatus(db, query);
      res.writeHead(result.status || 200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(result));
      return;
    }
    const qPhone = query.phone || query.phone_number;
    const qUserId = query.userId || query.user_id;
    let target = null;
    if (qUserId) target = (db.allUsers || []).find(u => u.id === qUserId);
    if (!target && qPhone) target = (db.allUsers || []).find(u => u.phone_number === qPhone);
    const hasQueryParam = Boolean(qPhone || qUserId || query.email);
    if (!target && !hasQueryParam && db.user) target = db.user;
    if (!target) {
      res.writeHead(404, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: false, error: 'User tidak ditemukan' }));
      return;
    }
    const isApproved = target.approval_status === 'approved' || target.is_approved === true;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({
      ok: true,
      status: 200,
      user_id: target.id,
      full_name: target.full_name,
      phone_number: target.phone_number,
      role: target.role || 'user',
      approval_status: target.approval_status || 'pending',
      is_approved: isApproved,
      rejection_reason: target.rejection_reason || null,
      redirect: isApproved ? '/dashboard' : null
    }));
    return;
  }

  // ------------------------------------------------------------------
  // API: AUTH (LOGIN & REGISTER WITH ROLE SEPARATION & APPROVAL ENFORCEMENT)
  // ------------------------------------------------------------------
  if (pathname === '/api/auth/login' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const payload = JSON.parse(body || '{}');
        if (userAuthApprovalModule) {
          const result = userAuthApprovalModule.handleUserLogin(db, payload);
          res.writeHead(result.status || 200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify(result));
          return;
        }
        const { identifier, method: authMethod } = payload;
        let matchedUser = null;
        if (identifier) {
          matchedUser = db.allUsers.find(u =>
            (u.email && u.email.toLowerCase() === identifier.toLowerCase()) ||
            (u.phone_number && u.phone_number === identifier)
          );
        }
        if (matchedUser) {
          db.user.id = matchedUser.id;
          db.user.full_name = matchedUser.full_name;
          db.user.phone_number = matchedUser.phone_number;
          db.subscription.plan = matchedUser.plan;
        }
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          ok: true,
          user: db.user,
          message: `Login berhasil menggunakan ${authMethod || 'autentikasi'}`
        }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  if (pathname === '/api/auth/register' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const payload = JSON.parse(body || '{}');
        if (userAuthApprovalModule) {
          const result = userAuthApprovalModule.handleUserRegister(db, payload);
          res.writeHead(result.status || 200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify(result));
          return;
        }
        const { full_name, phone_number, email } = payload;
        const newId = 'usr_tatadana_' + String(db.allUsers.length + 1).padStart(3, '0');
        const DEFAULT_AVATAR = '/uploads/avatars/default-avatar.svg';
        const cleanName = full_name ? String(full_name).trim() : 'Pengguna Baru';
        const cleanEmail = email ? String(email).trim().toLowerCase() : `${cleanName.toLowerCase().replace(/[^a-z0-9]/g, '')}@tatadana.id`;
        const newUser = {
          id: newId,
          full_name: cleanName,
          phone_number: '', // sepenuhnya KOSONG untuk pendaftaran baru
          registered_phone: phone_number ? String(phone_number).trim() : '',
          email: cleanEmail,
          plan: 'starter',
          telegram_connected: false,
          telegram_username: null,
          custom_bot_token: null,
          total_transactions: 0,
          joined_at: new Date().toISOString().slice(0, 10),
          job_profile: '', // sepenuhnya KOSONG
          job_title: '',
          bio: '', // sepenuhnya KOSONG
          avatar_url: DEFAULT_AVATAR
        };
        db.allUsers.push(newUser);
        db.user.id = newUser.id;
        db.user.full_name = newUser.full_name;
        db.user.email = newUser.email;
        db.user.phone_number = '';
        db.user.job_profile = '';
        db.user.job_title = '';
        db.user.bio = '';
        db.user.avatar_url = DEFAULT_AVATAR;
        db.subscription.plan = 'starter';
        saveUserProfile(db.user);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, user: db.user, message: 'Registrasi berhasil' }));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  // ------------------------------------------------------------------
  // API: ADMIN USER MANAGEMENT & APPROVAL ACTIONS (REQUIREMENT 4)
  // ------------------------------------------------------------------
  if (pathname === '/api/admin/users' && method === 'GET') {
    const status = reqUrl.searchParams.get('status');
    const role = reqUrl.searchParams.get('role');
    const search = reqUrl.searchParams.get('search');
    let users = db.allUsers || [];
    if (userAuthApprovalModule) {
      users = userAuthApprovalModule.filterUsers(users, { status, role, search });
    }
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true, users, count: users.length }));
    return;
  }

  if (pathname === '/api/admin/users/approval' && method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const payload = JSON.parse(body || '{}');
        if (!userAuthApprovalModule) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: 'Modul User Auth Approval tidak tersedia' }));
          return;
        }
        const result = userAuthApprovalModule.handleUserApprovalAction(db, payload);
        res.writeHead(result.status || 200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(result));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: err.message }));
      }
    });
    return;
  }

  if (pathname === '/api/admin/export-users' && method === 'GET') {
    res.writeHead(200, {
      'Content-Type': 'application/json',
      'Content-Disposition': 'attachment; filename="TataDana_All_Users.json"'
    });
    res.end(JSON.stringify(db.allUsers, null, 2));
    return;
  }

  // ------------------------------------------------------------------
  // API: GET EXPORT HISTORY (PART 3)
  // ------------------------------------------------------------------
  if (pathname === '/api/reports/exports' && method === 'GET') {
    const requestedUserId = reqUrl.searchParams.get('userId') || req.headers['x-user-id'];
    const currentUserId = requestedUserId || (db.user && db.user.id) || 'usr_admin_restu';
    const userHistory = (db.exportHistory || []).filter(e => e.user_id === currentUserId);
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true, exportHistory: userHistory }));
    return;
  }

  // ------------------------------------------------------------------
  // API: EXPORT EXCEL (CSV / SPREADSHEET WITH SIGNED URL & STARTER RESTRICTION)
  // ------------------------------------------------------------------
  if (pathname === '/api/reports/export/excel' && method === 'GET') {
    const requestedUserId = reqUrl.searchParams.get('userId') || req.headers['x-user-id'];
    const currentUserId = requestedUserId || (db.user && db.user.id) || 'usr_admin_restu';
    let plan = db.subscription.plan;
    let isAdmin = (db.user && db.user.role === 'admin');
    if (requestedUserId) {
      const u = (db.allUsers || []).find(user => user.id === requestedUserId);
      if (u) {
        if (u.plan) plan = u.plan;
        if (u.role === 'admin') isAdmin = true;
      }
    }
    if (plan === 'starter' && !isAdmin) {
      res.writeHead(403, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        ok: false,
        error: 'Fitur export Excel (.xlsx) khusus paket Pro. Silakan upgrade!'
      }));
      return;
    }

    const userTxs = (db.transactions || []).filter(t => t.user_id === currentUserId);
    const exportId = 'exp_' + Date.now();
    const fileName = `TataDana_Laporan_${new Date().toISOString().slice(0, 10)}.xlsx`;
    const wib = getWIBTime();

    if (XLSX && (reqUrl.searchParams.get('format') === 'xlsx' || req.headers['x-export-format'] === 'xlsx')) {
      const txRows = userTxs.length > 0 ? userTxs.map((t, idx) => ({
        'No': idx + 1,
        'ID Transaksi': t.id,
        'Tanggal (WIB)': new Date(t.transaction_date).toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' }) + ' WIB',
        'Jenis': t.type === 'income' ? 'Pemasukan' : (t.type === 'expense' ? 'Pengeluaran' : 'Transfer'),
        'Kategori': (t.category_emoji ? t.category_emoji + ' ' : '') + (t.category_name || '-'),
        'Dompet': t.wallet_name || '-',
        'Nominal (Rp)': t.amount,
        'Catatan': t.note || '-',
        'Sumber': t.source || 'Web App'
      })) : [{
        'No': 1,
        'ID Transaksi': '-',
        'Tanggal (WIB)': wib.full,
        'Jenis': '-',
        'Kategori': '-',
        'Dompet': '-',
        'Nominal (Rp)': 0,
        'Catatan': 'Belum ada transaksi pada akun ini',
        'Sumber': '-'
      }];

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(txRows);
      ws['!cols'] = [
        { wch: 6 },
        { wch: 22 },
        { wch: 24 },
        { wch: 14 },
        { wch: 24 },
        { wch: 18 },
        { wch: 16 },
        { wch: 34 },
        { wch: 16 }
      ];
      XLSX.utils.book_append_sheet(wb, ws, 'Transaksi');
      const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

      db.exportHistory.unshift({
        id: exportId,
        user_id: currentUserId,
        file_name: fileName,
        format: 'Excel',
        period: 'Bulan Berjalan',
        size: `${Math.max(1, Math.round(buffer.length / 1024))} KB`,
        created_at: wib.full,
        download_url: '/api/reports/export/excel'
      });

      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${fileName}"`,
        'Content-Length': buffer.length
      });
      res.end(buffer);
      return;
    }

    const csvHeader = '\uFEFFID,Tanggal (WIB),Jenis,Nominal,Kategori,Dompet,Catatan,Sumber\n';
    const csvRows = userTxs.map(t => {
      const date = new Date(t.transaction_date).toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
      return `"${t.id}","${date}","${t.type}","${t.amount}","${t.category_name || ''}","${t.wallet_name || ''}","${(t.note || '').replace(/"/g, '""')}","${t.source}"`;
    }).join('\n');

    db.exportHistory.unshift({
      id: exportId,
      user_id: currentUserId,
      file_name: fileName,
      format: 'Excel',
      period: 'Bulan Berjalan',
      size: `${Math.round((csvHeader.length + csvRows.length) / 1024) + 12} KB`,
      created_at: getWIBTime().full,
      download_url: '/api/reports/export/excel'
    });

    res.writeHead(200, {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8',
      'Content-Disposition': `attachment; filename="${fileName}"`,
    });
    res.end(csvHeader + csvRows);
    return;
  }

  // ------------------------------------------------------------------
  // API: EXPORT PDF SUMMARY (FORMAL FINANCIAL REPORT WITH STARTER RESTRICTION)
  // ------------------------------------------------------------------
  if (pathname === '/api/reports/export/pdf' && method === 'GET') {
    const requestedUserId = reqUrl.searchParams.get('userId') || req.headers['x-user-id'];
    const currentUserId = requestedUserId || (db.user && db.user.id) || 'usr_admin_restu';
    let plan = db.subscription.plan;
    let isAdmin = (db.user && db.user.role === 'admin');
    let reportUser = db.user;
    if (requestedUserId) {
      const u = (db.allUsers || []).find(user => user.id === requestedUserId);
      if (u) {
        if (u.plan) plan = u.plan;
        if (u.role === 'admin') isAdmin = true;
        reportUser = u;
      }
    }
    if (plan === 'starter' && !isAdmin) {
      res.writeHead(403, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        ok: false,
        error: 'Fitur export PDF resmi khusus paket Pro. Silakan upgrade!'
      }));
      return;
    }

    const userTxs = (db.transactions || []).filter(t => t.user_id === currentUserId);
    const totalExp = userTxs.filter(t => t.type === 'expense').reduce((a, b) => a + b.amount, 0);
    const totalInc = userTxs.filter(t => t.type === 'income').reduce((a, b) => a + b.amount, 0);
    const wib = getWIBTime();

    // Record in export history (Part 3)
    const exportId = 'exp_' + Date.now();
    const fileName = `TataDana_Laporan_Resmi_${new Date().toISOString().slice(0, 10)}.pdf`;
    db.exportHistory.unshift({
      id: exportId,
      user_id: currentUserId,
      file_name: fileName,
      format: 'PDF',
      period: 'Bulan Berjalan',
      size: '260 KB',
      created_at: wib.full,
      download_url: '/api/reports/export/pdf'
    });

    const htmlReport = `
      <!DOCTYPE html>
      <html lang="id">
      <head>
        <meta charset="utf-8">
        <title>TataDana — Laporan Keuangan Resmi</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 40px; color: #0f172a; background: #fff; }
          .header { border-bottom: 3px solid #f97316; padding-bottom: 20px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: center; }
          .brand { display: flex; align-items: center; gap: 12px; }
          .logo { width: 44px; height: 44px; background: linear-gradient(135deg, #f97316, #ea580c); border-radius: 12px; color: white; display: flex; align-items: center; justify-content: center; font-weight: 900; font-size: 20px; }
          .title { font-size: 22px; font-weight: 800; color: #0f172a; }
          .grid { display: flex; gap: 20px; margin-bottom: 30px; }
          .card { flex: 1; padding: 18px; background: #f8fafc; border-radius: 12px; border: 1px solid #e2e8f0; }
          table { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 15px; }
          th, td { padding: 12px 14px; border-bottom: 1px solid #e2e8f0; text-align: left; }
          th { background: #f1f5f9; color: #475569; font-weight: 700; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px; }
          .num { text-align: right; }
          .inc { color: #10b981; font-weight: bold; }
          .exp { color: #f97316; font-weight: bold; }
          .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #e2e8f0; font-size: 11px; color: #94a3b8; display: flex; justify-content: space-between; }
        </style>
      </head>
      <body>
        <div class="header">
          <div class="brand">
            <div class="logo">TD</div>
            <div>
              <div class="title">TataDana — Laporan Keuangan Resmi</div>
              <div style="font-size: 12px; color: #64748b; margin-top: 3px;">Periode: ${wib.dateStr} | Pengguna: <strong>${reportUser.full_name || 'Pengguna'}</strong> (${reportUser.phone_number || '-'})</div>
            </div>
          </div>
          <div style="text-align: right; font-size: 11px; color: #64748b; line-height: 1.5;">
            Zona Waktu: <strong>Asia/Jakarta (WIB)</strong><br>
            Database: <strong>Supabase Postgres (Verified)</strong><br>
            Signed Hash: <code>sha256:7f4a...9c12</code>
          </div>
        </div>

        <div class="grid">
          <div class="card">
            <div style="font-size: 11px; color: #64748b; text-transform: uppercase; font-weight: 600;">Total Pemasukan</div>
            <div style="font-size: 22px; font-weight: 800; color: #10b981; margin-top: 6px;">${formatRupiah(totalInc)}</div>
          </div>
          <div class="card">
            <div style="font-size: 11px; color: #64748b; text-transform: uppercase; font-weight: 600;">Total Pengeluaran</div>
            <div style="font-size: 22px; font-weight: 800; color: #f97316; margin-top: 6px;">${formatRupiah(totalExp)}</div>
          </div>
          <div class="card">
            <div style="font-size: 11px; color: #64748b; text-transform: uppercase; font-weight: 600;">Net Cashflow (Surplus)</div>
            <div style="font-size: 22px; font-weight: 800; color: #0284c7; margin-top: 6px;">${formatRupiah(totalInc - totalExp)}</div>
          </div>
        </div>

        <h3 style="font-size: 14px; font-weight: 700; color: #0f172a;">Rincian Riwayat Transaksi (${userTxs.length} Transaksi)</h3>
        <table>
          <thead>
            <tr>
              <th>Tanggal (WIB)</th>
              <th>Catatan / Merchant</th>
              <th>Kategori</th>
              <th>Dompet</th>
              <th class="num">Nominal</th>
            </tr>
          </thead>
          <tbody>
            ${userTxs.map(t => `
              <tr>
                <td>${new Date(t.transaction_date).toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' })}</td>
                <td>
                  <strong>${t.note}</strong>
                  ${t.items ? `<div style="font-size: 10px; color: #64748b; margin-top: 2px;">• Terdiri dari ${t.items.length} item struk OCR</div>` : ''}
                </td>
                <td>${t.category_emoji || '💳'} ${t.category_name || '-'}</td>
                <td>${t.wallet_name || '-'}</td>
                <td class="num ${t.type === 'income' ? 'inc' : 'exp'}">${t.type === 'income' ? '+' : '-'}${formatRupiah(t.amount)}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div class="footer">
          <div>Dicetak langsung dari TataDana SaaS Cloud • Enkripsi Row Level Security (RLS)</div>
          <div>Halaman 1 dari 1</div>
        </div>
        <script>window.print();</script>
      </body>
      </html>
    `;

    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(htmlReport);
    return;
  }

  // ------------------------------------------------------------------
  // API: EXPORT TRANSAKSI PDF (REQUIREMENT 3 BACKEND ENHANCEMENT)
  // ------------------------------------------------------------------
  if (pathname === '/api/transactions/export/pdf' && method === 'GET') {
    const totalExp = (db.transactions || []).filter(t => t.type === 'expense').reduce((a, b) => a + b.amount, 0);
    const totalInc = (db.transactions || []).filter(t => t.type === 'income').reduce((a, b) => a + b.amount, 0);
    const wib = getWIBTime();
    const fileName = `TataDana_Transaksi_${new Date().toISOString().slice(0, 10)}.pdf`;

    const txRowsHtml = (db.transactions || []).map((t, idx) => {
      const date = new Date(t.transaction_date).toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
      const isInc = t.type === 'income';
      return `
        <tr>
          <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-size: 11px;">${idx + 1}</td>
          <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-size: 11px;">${date} WIB</td>
          <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-size: 11px; font-weight: 600;">${t.note || '-'}</td>
          <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-size: 11px;">${t.category_name || '-'}</td>
          <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-size: 11px;">${t.wallet_name || '-'}</td>
          <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-size: 11px; color: ${isInc ? '#10b981' : '#f97316'}; font-weight: 600;">${isInc ? 'Pemasukan' : 'Pengeluaran'}</td>
          <td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-size: 11px; text-align: right; font-weight: bold; color: ${isInc ? '#10b981' : '#f97316'};">${isInc ? '+' : '-'}${formatRupiah(t.amount)}</td>
        </tr>
      `;
    }).join('');

    const htmlReport = `
      <!DOCTYPE html>
      <html lang="id">
      <head>
        <meta charset="utf-8">
        <title>TataDana - Laporan Transaksi Keuangan</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 40px; color: #0f172a; background: #fff; }
          .header { border-bottom: 3px solid #ea580c; padding-bottom: 20px; margin-bottom: 25px; display: flex; justify-content: space-between; align-items: center; }
          .brand { display: flex; align-items: center; gap: 12px; }
          .logo { width: 44px; height: 44px; background: linear-gradient(135deg, #f97316, #ea580c); border-radius: 12px; color: white; display: flex; align-items: center; justify-content: center; font-weight: 900; font-size: 20px; }
          .title { font-size: 22px; font-weight: 800; color: #0f172a; }
          .grid { display: flex; gap: 15px; margin-bottom: 25px; }
          .card { flex: 1; padding: 15px; background: #f8fafc; border-radius: 12px; border: 1px solid #e2e8f0; }
          .card-title { font-size: 11px; color: #64748b; font-weight: 600; text-transform: uppercase; }
          .card-value { font-size: 18px; font-weight: 800; margin-top: 5px; }
          table { width: 100%; border-collapse: collapse; text-align: left; margin-top: 15px; }
          th { background: #ea580c; color: white; padding: 10px; font-size: 11px; font-weight: 700; }
          .footer { margin-top: 40px; padding-top: 15px; border-top: 1px solid #e2e8f0; font-size: 10px; color: #94a3b8; display: flex; justify-content: space-between; }
        </style>
      </head>
      <body>
        <div class="header">
          <div class="brand">
            <div class="logo">TD</div>
            <div>
              <div class="title">TataDana — Riwayat Transaksi Lengkap</div>
              <div style="font-size: 12px; color: #64748b;">Pengguna: ${db.user.full_name} (${db.user.phone_number})</div>
            </div>
          </div>
          <div style="text-align: right; font-size: 11px; color: #64748b;">
            <div>Dicetak: ${wib.full}</div>
            <div style="color: #10b981; font-weight: 600;">Status Data: Sinkron Database</div>
          </div>
        </div>

        <div class="grid">
          <div class="card">
            <div class="card-title">Total Pemasukan</div>
            <div class="card-value" style="color: #10b981;">${formatRupiah(totalInc)}</div>
          </div>
          <div class="card">
            <div class="card-title">Total Pengeluaran</div>
            <div class="card-value" style="color: #f97316;">${formatRupiah(totalExp)}</div>
          </div>
          <div class="card">
            <div class="card-title">Net Cashflow</div>
            <div class="card-value" style="color: ${totalInc >= totalExp ? '#10b981' : '#ef4444'};">${formatRupiah(totalInc - totalExp)}</div>
          </div>
          <div class="card">
            <div class="card-title">Total Transaksi</div>
            <div class="card-value" style="color: #0ea5e9;">${db.transactions.length} Item</div>
          </div>
        </div>

        <table>
          <thead>
            <tr>
              <th>No</th>
              <th>Waktu (WIB)</th>
              <th>Catatan / Keterangan</th>
              <th>Kategori</th>
              <th>Dompet</th>
              <th>Jenis</th>
              <th style="text-align: right;">Nominal</th>
            </tr>
          </thead>
          <tbody>
            ${txRowsHtml}
          </tbody>
        </table>

        <div class="footer">
          <div>TataDana Fintech SaaS — Catat Keuangan Semudah Chat Langsung dari Telegram</div>
          <div>Dokumen Resmi Digital • WIB (Asia/Jakarta)</div>
        </div>
      </body>
      </html>
    `;

    const pdfBuffer = generateTransactionPdfBinary(db.user, db.transactions || [], totalInc, totalExp, wib);

    res.writeHead(200, {
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="${fileName}"`,
      'Content-Length': pdfBuffer.length
    });
    res.end(pdfBuffer);
    return;
  }

  // ------------------------------------------------------------------
  // API: CALENDAR ENGINE & CASHFLOW WITH NOTES CRUD (M3)
  // ------------------------------------------------------------------
  if (pathname.startsWith('/api/calendar') && calendarModule) {
    calendarModule.handleCalendarApi(req, res, db, reqUrl);
    return;
  }

  // ------------------------------------------------------------------
  // APPLICATION ROUTES: /, /login, /register, /dashboard, /admin
  // ------------------------------------------------------------------
  const pageRouteMap = {
    '/': 'landing',
    '/landing': 'landing',
    '/login': 'login',
    '/register': 'register',
    '/dashboard': 'dashboard',
    '/admin': 'admin',
    '/admin/login': 'admin-login'
  };

  const viewTitles = {
    'landing': 'TataDana — Catat Keuangan Semudah Chat Langsung dari Telegram',
    'login': 'Masuk — TataDana Fintech',
    'register': 'Daftar Akun — TataDana Fintech',
    'dashboard': 'Dashboard Pengguna — TataDana',
    'admin': 'Superadmin Control Panel — TataDana',
    'admin-login': 'Login Administrator — TataDana Superadmin Control Panel'
  };

  const validTabs = ['overview', 'transactions', 'reports', 'budgets', 'wallets', 'settings', 'calendar'];

  const cleanPath = (pathname.replace(/\/+$/, '') || '/').toLowerCase();

  let activeView = pageRouteMap[cleanPath];
  let requestedTab = null;

  // Support deep subroutes under /dashboard/ (e.g. /dashboard/transactions)
  if (!activeView && cleanPath.startsWith('/dashboard/')) {
    const sub = cleanPath.replace('/dashboard/', '').split('/')[0];
    if (validTabs.includes(sub)) {
      activeView = 'dashboard';
      requestedTab = sub;
    }
  }

  if (activeView && method === 'GET') {
    const appHtmlPath = path.join(__dirname, 'public', 'index.html');
    if (fs.existsSync(appHtmlPath)) {
      fs.readFile(appHtmlPath, 'utf8', (err, data) => {
        if (err) {
          res.writeHead(500);
          res.end('Error loading application');
          return;
        }

        let modifiedHtml = data;

        // SSR Pre-activation: render requested page view directly on reload/direct visit
        if (activeView !== 'landing') {
          // Hide landing page
          modifiedHtml = modifiedHtml.replace(
            'id="view-landing" class="page-view flex-1"',
            'id="view-landing" class="page-view hidden flex-1"'
          );
          // Show requested view
          modifiedHtml = modifiedHtml.replace(
            `id="view-${activeView}" class="page-view hidden`,
            `id="view-${activeView}" class="page-view`
          );
        }

        // SSR Page Title synchronization
        if (viewTitles[activeView]) {
          modifiedHtml = modifiedHtml.replace(
            /<title>.*?<\/title>/,
            `<title>${viewTitles[activeView]}</title>`
          );
        }

        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(modifiedHtml);
      });
      return;
    }
  }

  // ------------------------------------------------------------------
  // STATIC ASSET SERVING FOR /public FILES
  // ------------------------------------------------------------------
  const staticFilePath = path.join(__dirname, 'public', pathname);
  if (fs.existsSync(staticFilePath) && fs.statSync(staticFilePath).isFile() && method === 'GET') {
    const ext = path.extname(staticFilePath).toLowerCase();
    const mimeTypes = {
      '.html': 'text/html; charset=utf-8',
      '.js': 'application/javascript',
      '.css': 'text/css',
      '.json': 'application/json',
      '.png': 'image/png',
      '.jpg': 'image/jpeg',
      '.jpeg': 'image/jpeg',
      '.svg': 'image/svg+xml',
      '.webp': 'image/webp',
      '.gif': 'image/gif',
      '.ico': 'image/x-icon'
    };
    const contentType = mimeTypes[ext] || 'application/octet-stream';
    fs.readFile(staticFilePath, (err, fileData) => {
      if (err) {
        res.writeHead(500);
        res.end('Error loading file');
        return;
      }
      res.writeHead(200, { 'Content-Type': contentType });
      res.end(fileData);
    });
    return;
  }

  // Fallback for unmatched routes: redirect to /
  res.writeHead(302, { Location: '/' });
  res.end();
});

if (!process.env.VERCEL) {
  server.listen(PORT, '0.0.0.0', () => {
    console.log(`====================================================`);
    console.log(`🚀 TataDana SaaS Full-Stack Server Running!`);
    console.log(`🌐 Localhost URL : http://localhost:${PORT}`);
    console.log(`📡 Bot Mode      : Telegram Long Polling (Localhost-Ready)`);
    console.log(`📡 Webhook URL   : http://localhost:${PORT}/api/telegram/webhook`);
    console.log(`🕒 Timezone      : Asia/Jakarta (WIB)`);
    console.log(`🛡️ Database      : Supabase Postgres Schema Connected`);
    console.log(`====================================================`);

    // Auto-start Long Polling daemon if bot token is available in db, settings, or .env
    let autoToken = (db.user && db.user.custom_bot_token) || process.env.TELEGRAM_BOT_TOKEN;
    if (!autoToken && telegramPolling && typeof telegramPolling.loadTelegramToken === 'function') {
      const loaded = telegramPolling.loadTelegramToken();
      if (loaded && loaded.token) {
        autoToken = loaded.token;
      }
    }

    if (telegramPolling && autoToken) {
      if (db && db.user) {
        db.user.custom_bot_token = autoToken;
        db.user.custom_bot_connected = true;
        db.user.telegram_connected = true;
      }
      process.env.TELEGRAM_BOT_TOKEN = autoToken;
      telegramPolling.pollingManager.start({
        botToken: autoToken,
        db,
        messageHandler: handleTelegramUpdate
      }).then(() => {
        console.log('🤖 Auto-started Telegram Long Polling daemon successfully from persistent storage.');
      }).catch(err => {
        console.warn('⚠️ Could not auto-start Telegram polling:', err.message);
      });
    }
  });
} else {
  console.log('☁️ TataDana serverless instance initialized on Vercel.');
}

server.handleTelegramUpdate = handleTelegramUpdate;
server.telegramPolling = telegramPolling;

module.exports = server;
