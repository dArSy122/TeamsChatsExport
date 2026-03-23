export const LANG_STORAGE_KEY = "teams_export_lang";

export const SUPPORTED_LANGUAGES = [
  { code: "bg", label: "Български" },
  { code: "en", label: "English" },
  { code: "fr", label: "Français" },
];

const TRANSLATIONS = {
  bg: {
    appTitle: "Teams Chat Export Portal",
    loginTag: "Microsoft Teams",
    loginTitle: "Teams Chat Export Portal",
    loginText: "Защитен вътрешен портал за експортиране на Microsoft Teams чатове в напълно офлайн HTML архиви.",
    signInWithMicrosoft: "Вход с Microsoft",

    brandSub: "Избор и експорт на Teams чатове в единични HTML архиви за локално използване.",
    notSignedIn: "Не е влязъл",
    signedInAs: "Влязъл: {name}",

    step1Title: "Влез в портала",
    step1Text: "Използвай Microsoft акаунта си, за да се отвори достъп до наличните за теб Teams чатове.",
    step2Title: "Зареди чатовете",
    step2Text: "Натисни „Зареди чатове“, за да се изтегли пълният списък с достъпните 1:1 и групови чатове.",
    step3Title: "Избери какво да архивираш",
    step3Text: "Маркирай един, няколко или всички чатове, които искаш да включиш в export queue.",
    step4Title: "Стартирай експорт",
    step4Text: "Пусни „Експорт на избраните“, „Добави всички в опашката“ или „Повтори неуспешните“ според текущия статус.",
    step5Title: "Разреши свалянията",
    step5Text: "При първия експорт браузърът може да поиска разрешение за сваляне на множество файлове. Натисни Allow, за да се запазят всички архиви автоматично.",

    actionsTitle: "Основни действия",
    actionsText: "Използвай действията вдясно, за да заредиш списъка и да стартираш export queue.",
    loadChats: "Зареди чатове",
    exportSelected: "Експорт на избраните",
    queueAll: "Добави всички в опашката",
    retryFailedOnly: "Повтори неуспешните",

    chatSelectionTitle: "Избор на чатове",
    chatSelectionDesc: "Избери чатове от списъка, филтрирай по заглавие и следи статуса на queue обработката.",
    totalChats: "Общо чатове",
    selectedChats: "Избрани",
    searchPlaceholder: "Търси по тема, участници или видимо заглавие...",

    queueProgress: "Прогрес на опашката: {done} done • {running} running • {failed} failed • {total} total",
    queuePreparing: "Подготовка на опашката… {processed} / {total}",
    queueExporting: "Експорт {current} / {total}",
    queueFinished: "Опашката завърши {total} / {total}",

    statusQueued: "На опашка",
    statusRunning: "Обработва се",
    statusDone: "Готово",
    statusFailed: "Грешка",

    chatStarted: "Started",
    lastMessage: "Last message",
    queueLabel: "Queue",

    reason: "Причина",
    stage: "Етап",
    code: "Код",
    unknownError: "Непозната грешка",

    alreadySignedIn: "Вече си влязъл.",
    loginInProgress: "Входът вече е в ход. Изчакай redirect обработката да приключи.",
    loginRedirectStart: "Стартира Microsoft login redirect…",
    loginButtonClicked: "Натиснат е бутонът за вход.",
    loginFailed: "Входът не бе успешен",
    loginError: "Грешка при вход",
    pleaseLoginFirst: "Моля, първо влез в системата.",
    loadChatsFirst: "Първо зареди чатовете.",
    loadingChats: "Зареждане на чатове…",
    chatsLoaded: "Чатовете са заредени успешно: {count}",
    loadChatsFailed: "Зареждането на чатовете не бе успешно",

    selectAtLeastOneChat: "Моля, избери поне един чат.",
    noQueuedChats: "Няма чатове на опашка за експорт.",
    noFailedChats: "Няма неуспешни чатове за повторение.",

    queuePrepared: "Опашката е подготвена: добавени {added} / избрани {totalSelected}.",
    queueAllPrepared: "Подготвена е опашка за всички: добавени {added} / общо чатове {total}.",
    retryFailedPrepared: "Повторение само на неуспешните: върнати в опашката {count} чата.",

    queueStart: "Старт на опашката: {count} чат(а).",
    exportFolderSubfolder: "Експортите ще бъдат записани в подпапка: {folder}",
    exportFolderSelected: "Папката за експорт е избрана успешно.",
    exportFolderCancelled: "Изборът на папка за експорт беше отказан.",
    exportFolderFailed: "Изборът на папка за експорт не бе успешен: {message}",
    fsApiFallback: "File System Access API не е наличен. Преминава се към browser downloads.",
    allowMultipleDownloads: "Забележка: браузърът може да поиска разрешение за множество автоматични сваляния за този сайт.",

    queueItemRunning: "ОБРАБОТКА: {title}",
    queueItemDone: "ГОТОВО: {title} -> {file} ({items} елемента, период {period})",
    queueItemFailed: "ГРЕШКА: {title} -> {details}",
    queueFinishedSummary: "Опашката приключи. Готови: {done}, Грешки: {failed}, Общо: {total}",

    readySignedInAs: "Готово. Влязъл като {name}",
    notSignedInUseButton: "Не си влязъл. Използвай бутона за вход с Microsoft.",
    initializingAuth: "Инициализиране на auth…",

    sessionExpired: "Сесията изтече: не беше стартиран експорт в рамките на 3 минути. Моля, влез отново.",
    reloginRedirectFailed: "Пренасочването за повторен вход не бе успешно: {message}",
    reloginFailed: "Повторният вход не бе успешен",

    msalCacheMismatchRetry: "Открит е MSAL redirect cache mismatch. Опитай входа още веднъж.",
    interactionAlreadyInProgress: "MSAL interaction вече е в ход: {status}",

    exportQueueFailed: "Export queue failed",
    queueAllFailed: "Queue all failed",
    retryFailedError: "Retry failed only error",
    initFailed: "Init failed",
    uiMissingElementIds: "Липсват UI element IDs:\n{ids}",

    attachmentFileFallback: "Файл",
    attachmentPreview: "Преглед",
    attachmentDownloadOffline: "Свали офлайн",
    attachmentMissingGeneric: "файлът не е вграден",
    attachmentMissingUnavailable: "файлът не е бил наличен за сваляне от чата",
    attachmentMissingDeclinedLarge: "файлът не е вграден (потребителят отказа голям файл)",
    attachmentMissingBlockedType: "файлът не е вграден (неподдържан тип файл)",
    attachmentMissingTooLarge: "файлът е твърде голям за вграждане",
    attachmentMissingTotalLimit: "надвишен е общият лимит за файлове",
    attachmentMissingUnresolved: "файлът не може да бъде извлечен от Teams",

    exportArchiveDefaultTitle: "Teams Chat Archive (OFFLINE)",
    exportItems: "Елементи",
    exportPeriod: "Период",
    exportGenerated: "Генериран",
    exportImages: "Снимки",
    exportFiles: "Файлове",
    exportEmbedded: "embedded",
    exportParticipants: "Участници",
    exportLatest: "Най-нови",
    exportParticipantsTitle: "Участници в чата",
    exportParticipantsSubtitle: "Хората, които са част от този разговор",
    exportParticipantsCount: "{count} participant{s}",
    exportYou: "Ти",
    exportReactions: "Реакции",
    exportReactionListTitle: "{emoji} {count} reaction{s} • {reactionType}",
    exportNoParticipantNames: "Няма налични имена на участници",
    exportFilePreview: "Преглед на файл",
    exportOfflinePreview: "Офлайн преглед",
    exportImagePreview: "Преглед на изображение",
    exportCloseHint: "Натисни Esc или извън прозореца, за да затвориш",
    exportUnsupportedPreviewTitle: "Прегледът не се поддържа офлайн за този файлов тип",
    exportUnsupportedPreviewText: "Файлът е вграден в архива и може да бъде свален локално, но този файлов тип не може да бъде визуализиран надеждно директно от браузъра в офлайн режим.",
    exportFileType: "Тип файл",
    exportDownloadOffline: "Свали офлайн",
    exportPreview: "Преглед",
    exportForwarded: "Препратено",
    exportNoContent: "Няма съдържание",
    exportQuote: "Цитат",
    exportView: "Виж",
    exportDeletedMessage: "Това съобщение е изтрито",
    exportSystemEventFallback: "[Системно събитие]",
  },

  en: {
    appTitle: "Teams Chat Export Portal",
    loginTag: "Microsoft Teams",
    loginTitle: "Teams Chat Export Portal",
    loginText: "Secure internal portal for exporting Microsoft Teams chats into fully offline HTML archives.",
    signInWithMicrosoft: "Sign in with Microsoft",

    brandSub: "Select and export Teams chats into standalone HTML archives for local use.",
    notSignedIn: "Not signed in",
    signedInAs: "Signed in: {name}",

    step1Title: "Sign in to the portal",
    step1Text: "Use your Microsoft account to unlock access to the Teams chats available to you.",
    step2Title: "Load chats",
    step2Text: "Click “Load chats” to retrieve the full list of available 1:1 and group chats.",
    step3Title: "Choose what to archive",
    step3Text: "Select one, multiple, or all chats that you want to add to the export queue.",
    step4Title: "Start export",
    step4Text: "Use “Export selected”, “Queue all”, or “Retry failed only” depending on the current status.",
    step5Title: "Allow downloads",
    step5Text: "On the first export, the browser may ask for permission to download multiple files. Click Allow so all archives can be saved automatically.",

    actionsTitle: "Main actions",
    actionsText: "Use the actions on the right to load the list and start the export queue.",
    loadChats: "Load chats",
    exportSelected: "Export selected",
    queueAll: "Queue all",
    retryFailedOnly: "Retry failed only",

    chatSelectionTitle: "Chat selection",
    chatSelectionDesc: "Choose chats from the list, filter by title, and monitor queue processing status.",
    totalChats: "Total chats",
    selectedChats: "Selected",
    searchPlaceholder: "Search by topic, participants or visible title...",

    queueProgress: "Queue progress: {done} done • {running} running • {failed} failed • {total} total",
    queuePreparing: "Preparing queue… {processed} / {total}",
    queueExporting: "Exporting {current} / {total}",
    queueFinished: "Queue finished {total} / {total}",

    statusQueued: "Queued",
    statusRunning: "Running",
    statusDone: "Done",
    statusFailed: "Failed",

    chatStarted: "Started",
    lastMessage: "Last message",
    queueLabel: "Queue",

    reason: "Reason",
    stage: "Stage",
    code: "Code",
    unknownError: "Unknown error",

    alreadySignedIn: "Already signed in.",
    loginInProgress: "Login already in progress. Please wait for redirect handling to finish.",
    loginRedirectStart: "Starting Microsoft login redirect…",
    loginButtonClicked: "Login button clicked.",
    loginFailed: "Login failed",
    loginError: "Login error",
    pleaseLoginFirst: "Please login first.",
    loadChatsFirst: "Load chats first.",
    loadingChats: "Loading chats…",
    chatsLoaded: "Chats loaded successfully: {count}",
    loadChatsFailed: "Load chats failed",

    selectAtLeastOneChat: "Please select at least one chat first.",
    noQueuedChats: "No queued chats to export.",
    noFailedChats: "There are no failed chats to retry.",

    queuePrepared: "Queue prepared: added {added} / selected {totalSelected}.",
    queueAllPrepared: "Queue ALL prepared: added {added} / total chats {total}.",
    retryFailedPrepared: "Retry FAILED only: re-queued {count} chat(s).",

    queueStart: "Queue start: {count} chat(s).",
    exportFolderSubfolder: "Exports will be saved in subfolder: {folder}",
    exportFolderSelected: "Export folder selected successfully.",
    exportFolderCancelled: "Export folder selection was cancelled.",
    exportFolderFailed: "Export folder selection failed: {message}",
    fsApiFallback: "File System Access API is not available. Falling back to browser downloads.",
    allowMultipleDownloads: "Note: browser may ask to allow multiple automatic downloads for this site.",

    queueItemRunning: "RUNNING: {title}",
    queueItemDone: "DONE: {title} -> {file} ({items} items, period {period})",
    queueItemFailed: "FAILED: {title} -> {details}",
    queueFinishedSummary: "Queue finished. Done: {done}, Failed: {failed}, Total: {total}",

    readySignedInAs: "Ready. Signed in as {name}",
    notSignedInUseButton: "Not signed in. Use the Microsoft sign-in button.",
    initializingAuth: "Initializing auth…",

    sessionExpired: "Session expired: no export was started within 3 minutes. Please sign in again.",
    reloginRedirectFailed: "Re-login redirect failed: {message}",
    reloginFailed: "Re-login failed",

    msalCacheMismatchRetry: "MSAL redirect cache mismatch detected. Retry the sign-in once.",
    interactionAlreadyInProgress: "MSAL interaction already in progress: {status}",

    exportQueueFailed: "Export queue failed",
    queueAllFailed: "Queue all failed",
    retryFailedError: "Retry failed only error",
    initFailed: "Init failed",
    uiMissingElementIds: "UI missing element IDs:\n{ids}",

    attachmentFileFallback: "File",
    attachmentPreview: "Preview",
    attachmentDownloadOffline: "Download offline",
    attachmentMissingGeneric: "file was not embedded",
    attachmentMissingUnavailable: "the file was not available for download from the chat",
    attachmentMissingDeclinedLarge: "file was not embedded (user declined large file)",
    attachmentMissingBlockedType: "file was not embedded (unsupported file type)",
    attachmentMissingTooLarge: "file is too large to embed",
    attachmentMissingTotalLimit: "total file limit was exceeded",
    attachmentMissingUnresolved: "file could not be extracted from Teams",

    exportArchiveDefaultTitle: "Teams Chat Archive (OFFLINE)",
    exportItems: "Items",
    exportPeriod: "Period",
    exportGenerated: "Generated",
    exportImages: "Images",
    exportFiles: "Files",
    exportEmbedded: "embedded",
    exportParticipants: "Participants",
    exportLatest: "Latest",
    exportParticipantsTitle: "Chat participants",
    exportParticipantsSubtitle: "People who are part of this conversation",
    exportParticipantsCount: "{count} participant{s}",
    exportYou: "You",
    exportReactions: "Reactions",
    exportReactionListTitle: "{emoji} {count} reaction{s} • {reactionType}",
    exportNoParticipantNames: "No participant names available",
    exportFilePreview: "File preview",
    exportOfflinePreview: "Offline preview",
    exportImagePreview: "Image preview",
    exportCloseHint: "Press Esc or click outside the window to close",
    exportUnsupportedPreviewTitle: "Offline preview is not supported for this file type",
    exportUnsupportedPreviewText: "The file is embedded in the archive and can be downloaded locally, but this file type cannot be reliably previewed directly by the browser in offline mode.",
    exportFileType: "File type",
    exportDownloadOffline: "Download offline",
    exportPreview: "Preview",
    exportForwarded: "Forwarded",
    exportNoContent: "No content",
    exportQuote: "Quote",
    exportView: "View",
    exportDeletedMessage: "This message has been deleted.",
    exportSystemEventFallback: "[System event]",
  },

  fr: {
    appTitle: "Teams Chat Export Portal",
    loginTag: "Microsoft Teams",
    loginTitle: "Teams Chat Export Portal",
    loginText: "Portail interne sécurisé pour exporter les conversations Microsoft Teams dans des archives HTML entièrement hors ligne.",
    signInWithMicrosoft: "Se connecter avec Microsoft",

    brandSub: "Sélectionnez et exportez les conversations Teams dans des archives HTML autonomes pour un usage local.",
    notSignedIn: "Non connecté",
    signedInAs: "Connecté : {name}",

    step1Title: "Connecte-toi au portail",
    step1Text: "Utilise ton compte Microsoft pour accéder aux conversations Teams qui te sont disponibles.",
    step2Title: "Charge les conversations",
    step2Text: "Clique sur « Charger les conversations » pour récupérer la liste complète des conversations 1:1 et de groupe.",
    step3Title: "Choisis quoi archiver",
    step3Text: "Sélectionne une, plusieurs ou toutes les conversations à ajouter à la file d’export.",
    step4Title: "Démarre l’export",
    step4Text: "Utilise « Exporter la sélection », « Tout mettre en file » ou « Réessayer les échecs » selon l’état actuel.",
    step5Title: "Autorise les téléchargements",
    step5Text: "Lors du premier export, le navigateur peut demander l’autorisation de télécharger plusieurs fichiers. Clique sur Allow pour que toutes les archives soient enregistrées automatiquement.",

    actionsTitle: "Actions principales",
    actionsText: "Utilise les actions à droite pour charger la liste et démarrer la file d’export.",
    loadChats: "Charger les conversations",
    exportSelected: "Exporter la sélection",
    queueAll: "Tout mettre en file",
    retryFailedOnly: "Réessayer les échecs",

    chatSelectionTitle: "Sélection des conversations",
    chatSelectionDesc: "Choisis des conversations dans la liste, filtre par titre et suis l’état du traitement de la file.",
    totalChats: "Conversations totales",
    selectedChats: "Sélectionnées",
    searchPlaceholder: "Rechercher par sujet, participants ou titre visible...",

    queueProgress: "Progression de la file : {done} done • {running} running • {failed} failed • {total} total",
    queuePreparing: "Préparation de la file… {processed} / {total}",
    queueExporting: "Export en cours {current} / {total}",
    queueFinished: "File terminée {total} / {total}",

    statusQueued: "En file",
    statusRunning: "En cours",
    statusDone: "Terminé",
    statusFailed: "Échec",

    chatStarted: "Started",
    lastMessage: "Last message",
    queueLabel: "Queue",

    reason: "Raison",
    stage: "Étape",
    code: "Code",
    unknownError: "Erreur inconnue",

    alreadySignedIn: "Déjà connecté.",
    loginInProgress: "La connexion est déjà en cours. Attends la fin du traitement de redirection.",
    loginRedirectStart: "Démarrage de la redirection de connexion Microsoft…",
    loginButtonClicked: "Le bouton de connexion a été cliqué.",
    loginFailed: "Échec de la connexion",
    loginError: "Erreur de connexion",
    pleaseLoginFirst: "Connecte-toi d’abord.",
    loadChatsFirst: "Charge d’abord les conversations.",
    loadingChats: "Chargement des conversations…",
    chatsLoaded: "Conversations chargées avec succès : {count}",
    loadChatsFailed: "Le chargement des conversations a échoué",

    selectAtLeastOneChat: "Sélectionne au moins une conversation.",
    noQueuedChats: "Aucune conversation en file pour l’export.",
    noFailedChats: "Aucun échec à réessayer.",

    queuePrepared: "File préparée : {added} ajoutée(s) / {totalSelected} sélectionnée(s).",
    queueAllPrepared: "File complète préparée : {added} ajoutée(s) / {total} conversations au total.",
    retryFailedPrepared: "Réessai des échecs uniquement : {count} conversation(s) remise(s) en file.",

    queueStart: "Démarrage de la file : {count} conversation(s).",
    exportFolderSubfolder: "Les exports seront enregistrés dans le sous-dossier : {folder}",
    exportFolderSelected: "Le dossier d’export a été sélectionné avec succès.",
    exportFolderCancelled: "La sélection du dossier d’export a été annulée.",
    exportFolderFailed: "La sélection du dossier d’export a échoué : {message}",
    fsApiFallback: "File System Access API n’est pas disponible. Basculement vers les téléchargements du navigateur.",
    allowMultipleDownloads: "Remarque : le navigateur peut demander l’autorisation pour plusieurs téléchargements automatiques sur ce site.",

    queueItemRunning: "EN COURS : {title}",
    queueItemDone: "TERMINÉ : {title} -> {file} ({items} éléments, période {period})",
    queueItemFailed: "ÉCHEC : {title} -> {details}",
    queueFinishedSummary: "File terminée. Terminés : {done}, Échecs : {failed}, Total : {total}",

    readySignedInAs: "Prêt. Connecté en tant que {name}",
    notSignedInUseButton: "Non connecté. Utilise le bouton de connexion Microsoft.",
    initializingAuth: "Initialisation de l’authentification…",

    sessionExpired: "Session expirée : aucun export n’a été démarré dans les 3 minutes. Veuillez vous reconnecter.",
    reloginRedirectFailed: "La redirection de reconnexion a échoué : {message}",
    reloginFailed: "La reconnexion a échoué",

    msalCacheMismatchRetry: "Un décalage du cache de redirection MSAL a été détecté. Réessaie la connexion.",
    interactionAlreadyInProgress: "Une interaction MSAL est déjà en cours : {status}",

    exportQueueFailed: "Échec de la file d’export",
    queueAllFailed: "Échec de la mise en file complète",
    retryFailedError: "Erreur lors du réessai des échecs",
    initFailed: "Échec de l’initialisation",
    uiMissingElementIds: "IDs d’éléments UI manquants :\n{ids}",

    attachmentFileFallback: "Fichier",
    attachmentPreview: "Aperçu",
    attachmentDownloadOffline: "Télécharger hors ligne",
    attachmentMissingGeneric: "le fichier n’a pas été intégré",
    attachmentMissingUnavailable: "le fichier n’était pas disponible au téléchargement depuis la conversation",
    attachmentMissingDeclinedLarge: "le fichier n’a pas été intégré (gros fichier refusé par l’utilisateur)",
    attachmentMissingBlockedType: "le fichier n’a pas été intégré (type de fichier non pris en charge)",
    attachmentMissingTooLarge: "le fichier est trop volumineux pour être intégré",
    attachmentMissingTotalLimit: "la limite totale des fichiers a été dépassée",
    attachmentMissingUnresolved: "le fichier n’a pas pu être extrait depuis Teams",

    exportArchiveDefaultTitle: "Archive de conversation Teams (HORS LIGNE)",
    exportItems: "Éléments",
    exportPeriod: "Période",
    exportGenerated: "Généré",
    exportImages: "Images",
    exportFiles: "Fichiers",
    exportEmbedded: "intégrés",
    exportParticipants: "Participants",
    exportLatest: "Derniers",
    exportParticipantsTitle: "Participants de la conversation",
    exportParticipantsSubtitle: "Personnes faisant partie de cette conversation",
    exportParticipantsCount: "{count} participant{s}",
    exportYou: "Vous",
    exportReactions: "Réactions",
    exportReactionListTitle: "{emoji} {count} réaction{s} • {reactionType}",
    exportNoParticipantNames: "Aucun nom de participant disponible",
    exportFilePreview: "Aperçu du fichier",
    exportOfflinePreview: "Aperçu hors ligne",
    exportImagePreview: "Aperçu de l’image",
    exportCloseHint: "Appuie sur Échap ou clique hors de la fenêtre pour fermer",
    exportUnsupportedPreviewTitle: "L’aperçu hors ligne n’est pas pris en charge pour ce type de fichier",
    exportUnsupportedPreviewText: "Le fichier est intégré dans l’archive et peut être téléchargé localement, mais ce type de fichier ne peut pas être prévisualisé de manière fiable directement dans le navigateur en mode hors ligne.",
    exportFileType: "Type de fichier",
    exportDownloadOffline: "Télécharger hors ligne",
    exportPreview: "Aperçu",
    exportForwarded: "Transféré",
    exportNoContent: "Pas de contenu",
    exportQuote: "Citation",
    exportView: "Voir",
    exportDeletedMessage: "Ce message a été supprimé.",
    exportSystemEventFallback: "[Événement système]",
  },
};

let currentLang = loadLanguage();

function normalizeLang(lang) {
  const raw = String(lang || "").trim().toLowerCase();
  return TRANSLATIONS[raw] ? raw : "en";
}

export function loadLanguage() {
  const raw = localStorage.getItem(LANG_STORAGE_KEY);
  return normalizeLang(raw || "bg");
}

export function getCurrentLanguage() {
  return currentLang;
}

export function setCurrentLanguage(lang) {
  currentLang = normalizeLang(lang);
  localStorage.setItem(LANG_STORAGE_KEY, currentLang);
  return currentLang;
}

export function getSupportedLanguages() {
  return SUPPORTED_LANGUAGES.slice();
}

function interpolate(text, vars = {}) {
  let out = String(text ?? "");
  for (const [name, value] of Object.entries(vars)) {
    out = out.replaceAll(`{${name}}`, String(value ?? ""));
  }
  return out;
}

export function tFor(lang, key, vars = {}) {
  const normalized = normalizeLang(lang);
  const dict = TRANSLATIONS[normalized] || TRANSLATIONS.en;
  const fallback = TRANSLATIONS.en || {};
  const text = dict[key] ?? fallback[key] ?? key;
  return interpolate(text, vars);
}

export function t(key, vars = {}) {
  return tFor(currentLang, key, vars);
}

export function applyI18nToDocument(root = document) {
  root.querySelectorAll("[data-i18n]").forEach((el) => {
    const key = el.getAttribute("data-i18n");
    if (!key) return;
    el.textContent = t(key);
  });

  root.querySelectorAll("[data-i18n-placeholder]").forEach((el) => {
    const key = el.getAttribute("data-i18n-placeholder");
    if (!key) return;
    el.setAttribute("placeholder", t(key));
  });

  root.querySelectorAll("[data-i18n-title]").forEach((el) => {
    const key = el.getAttribute("data-i18n-title");
    if (!key) return;
    el.setAttribute("title", t(key));
  });

  document.documentElement.lang = currentLang;
}