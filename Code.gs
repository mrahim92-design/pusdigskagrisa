/**
 * Pustaka Digital Premium - Google Apps Script Backend (Code.gs)
 * Developed for professional school/library digital reading systems.
 */

// Configuration Variables
// Replace these with your actual IDs when deploying, or leave blank to auto-create inside the active spreadsheet
var SPREADSHEET_ID = ""; 
var PDF_FOLDER_NAME = "Pustaka_PDF_Books";
var COVER_FOLDER_NAME = "Pustaka_Book_Covers";
var AVATAR_FOLDER_NAME = "Pustaka_User_Photos";

/**
 * Serves the HTML web app
 */
function doGet(e) {
  checkAndInitDatabase();
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('Pustaka Digital Premium')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Includes HTML templates inside others
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (err) {
    return `<!-- Error including ${filename}: ${err.message} -->`;
  }
}

/**
 * Accesses or opens the spreadsheet database
 */
function getDb() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    // If not running in a container-bound script, prompt or use current active
    var files = DriveApp.getFilesByName("Pustaka_Digital_Premium_DB");
    if (files.hasNext()) {
      return SpreadsheetApp.open(files.next());
    } else {
      var ss = SpreadsheetApp.create("Pustaka_Digital_Premium_DB");
      return ss;
    }
  }
}

/**
 * Accesses or creates folders in Google Drive for storage
 */
function getFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    var folder = DriveApp.createFolder(folderName);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return folder;
  }
}

/**
 * Automatically initializes sheets and tables with default headers and data if not present
 */
function checkAndInitDatabase() {
  var db = getDb();
  
  // Sheet 1: USERS
  var sheetUsers = db.getSheetByName("USERS");
  if (!sheetUsers) {
    sheetUsers = db.insertSheet("USERS");
    sheetUsers.appendRow(["ID_USER", "NAMA", "USERNAME", "PASSWORD", "ROLE", "KELAS", "FOTO", "STATUS", "TANGGAL_DAFTAR"]);
  }
  
  // If no users, create default admin
  if (sheetUsers.getLastRow() <= 1) {
    var salt = "pustaka_premium_salt";
    var defaultPass = hashPassword("admin123", salt);
    sheetUsers.appendRow([
      "USR-001",
      "Super Administrator",
      "admin",
      defaultPass,
      "Admin",
      "Staf IT",
      "", // default avatar empty
      "Aktif",
      new Date().toISOString()
    ]);
    
    // Create demo Pembaca user
    var demoPass = hashPassword("pembaca123", salt);
    sheetUsers.appendRow([
      "USR-002",
      "Pembaca Demo",
      "pembaca",
      demoPass,
      "Pembaca",
      "Kelas XII-IPA",
      "",
      "Aktif",
      new Date().toISOString()
    ]);
  }

  // Sheet 2: BUKU
  var sheetBuku = db.getSheetByName("BUKU");
  if (!sheetBuku) {
    sheetBuku = db.insertSheet("BUKU");
    sheetBuku.appendRow(["ID_BUKU", "JUDUL", "PENULIS", "PENERBIT", "TAHUN", "KATEGORI", "DESKRIPSI", "COVER", "LINK_PDF", "JUMLAH_DIBACA", "STATUS", "TANGGAL_UPLOAD"]);
    
    // Add Demo Book
    sheetBuku.appendRow([
      "BKP-001",
      "Pedoman Penggunaan Pustaka Digital Premium",
      "Tim Pengembang",
      "Pustaka Premium Press",
      "2026",
      "Petunjuk",
      "Panduan lengkap mengenai tata tertib dan cara memaksimalkan fitur Pustaka Digital Premium.",
      "https://images.unsplash.com/photo-1543002588-bfa74002ed7e?auto=format&fit=crop&q=80&w=400", // Unsplash premium book cover placeholder
      "https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf", // Standard testing dummy pdf
      "0",
      "Aktif",
      new Date().toISOString()
    ]);
  }

  // Sheet 3: KATEGORI
  var sheetKategori = db.getSheetByName("KATEGORI");
  if (!sheetKategori) {
    sheetKategori = db.insertSheet("KATEGORI");
    sheetKategori.appendRow(["ID_KATEGORI", "NAMA_KATEGORI", "ICON", "WARNA"]);
    
    // Add default categories
    sheetKategori.appendRow(["CAT-001", "Petunjuk", "fa-info-circle", "indigo"]);
    sheetKategori.appendRow(["CAT-002", "Sains", "fa-flask", "emerald"]);
    sheetKategori.appendRow(["CAT-003", "Teknologi", "fa-laptop-code", "blue"]);
    sheetKategori.appendRow(["CAT-004", "Sastra & Novel", "fa-book-open", "rose"]);
  }

  // Sheet 4: RIWAYAT_BACA
  var sheetRiwayat = db.getSheetByName("RIWAYAT_BACA");
  if (!sheetRiwayat) {
    sheetRiwayat = db.insertSheet("RIWAYAT_BACA");
    sheetRiwayat.appendRow(["ID", "USER", "ID_BUKU", "JUDUL", "WAKTU_BACA"]);
  }

  // Clean empty sheet created by default
  var defaultSheet = db.getSheetByName("Sheet1");
  if (defaultSheet) {
    try {
      db.deleteSheet(defaultSheet);
    } catch(err) {}
  }
}

/**
 * SHA-256 Password hashing with salt helper
 */
function hashPassword(password, salt) {
  if (!salt) salt = "pustaka_premium_salt";
  var combined = password + salt;
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, combined, Utilities.Charset.UTF_8);
  var hashStr = "";
  for (var i = 0; i < rawHash.length; i++) {
    var val = rawHash[i];
    if (val < 0) val += 256;
    var byteString = val.toString(16);
    if (byteString.length == 1) byteString = "0" + byteString;
    hashStr += byteString;
  }
  return hashStr;
}

/**
 * Helper to convert sheet into Array of Objects
 */
function getSheetDataAsObjects(sheetName) {
  var db = getDb();
  var sheet = db.getSheetByName(sheetName);
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return [];
  
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  
  return values.map(function(row) {
    var obj = {};
    headers.forEach(function(header, colIdx) {
      var val = row[colIdx];
      // Format Date values if encountered
      if (val instanceof Date) {
        val = val.toISOString();
      }
      obj[header] = val;
    });
    return obj;
  });
}

/**
 * Base64 file upload helper for Google Drive
 */
function uploadFileToDrive(base64Data, filename, folderName, mimeType) {
  try {
    var folder = getFolder(folderName);
    var bytes = Utilities.base64Decode(base64Data.split(",")[1]);
    var blob = Utilities.newBlob(bytes, mimeType, filename);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    throw new Error("Gagal mengunggah file " + filename + " ke Drive: " + err.message);
  }
}

// ==================== AUTHENTICATION API ====================

function loginUser(username, password) {
  try {
    checkAndInitDatabase();
    var users = getSheetDataAsObjects("USERS");
    var hashed = hashPassword(password, "pustaka_premium_salt");
    
    var matchedUser = users.find(function(u) {
      return u.USERNAME.toLowerCase() === username.toLowerCase() && u.PASSWORD === hashed;
    });
    
    if (!matchedUser) {
      return { success: false, message: "Username atau Password salah!" };
    }
    
    if (matchedUser.STATUS !== "Aktif") {
      return { success: false, message: "Akun Anda dinonaktifkan. Silakan hubungi Administrator!" };
    }
    
    // Do not return password hash to client
    delete matchedUser.PASSWORD;
    
    return { success: true, user: matchedUser };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function registerUser(userData) {
  try {
    var db = getDb();
    var sheet = db.getSheetByName("USERS");
    var users = getSheetDataAsObjects("USERS");
    
    // Check duplication
    var dup = users.some(function(u) {
      return u.USERNAME.toLowerCase() === userData.USERNAME.toLowerCase();
    });
    if (dup) {
      return { success: false, message: "Username sudah terdaftar!" };
    }
    
    // Generate new User ID
    var newIdNum = users.length + 1;
    var newId = "USR-" + String(newIdNum).padStart(3, '0');
    
    var salt = "pustaka_premium_salt";
    var hashed = hashPassword(userData.PASSWORD, salt);
    
    var photoUrl = "";
    if (userData.PHOTO_BASE64) {
      photoUrl = uploadFileToDrive(userData.PHOTO_BASE64, newId + "_avatar", AVATAR_FOLDER_NAME, userData.PHOTO_MIME || "image/png");
    }
    
    sheet.appendRow([
      newId,
      userData.NAMA,
      userData.USERNAME,
      hashed,
      userData.ROLE || "Pembaca",
      userData.KELAS || "Umum",
      photoUrl,
      "Aktif",
      new Date().toISOString()
    ]);
    
    return { success: true, message: "Pendaftaran berhasil!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function updateProfile(userId, updateData) {
  try {
    var db = getDb();
    var sheet = db.getSheetByName("USERS");
    var users = getSheetDataAsObjects("USERS");
    var rowIdx = users.findIndex(function(u) { return u.ID_USER === userId; });
    
    if (rowIdx === -1) {
      return { success: false, message: "User tidak ditemukan!" };
    }
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var actualRow = rowIdx + 2; // +1 header offset, +1 1-based indexing
    
    if (updateData.PHOTO_BASE64) {
      var photoUrl = uploadFileToDrive(updateData.PHOTO_BASE64, userId + "_avatar", AVATAR_FOLDER_NAME, updateData.PHOTO_MIME || "image/png");
      var photoCol = headers.indexOf("FOTO") + 1;
      sheet.getRange(actualRow, photoCol).setValue(photoUrl);
    }
    
    if (updateData.PASSWORD) {
      var hashed = hashPassword(updateData.PASSWORD, "pustaka_premium_salt");
      var passCol = headers.indexOf("PASSWORD") + 1;
      sheet.getRange(actualRow, passCol).setValue(hashed);
    }
    
    if (updateData.NAMA) {
      var namaCol = headers.indexOf("NAMA") + 1;
      sheet.getRange(actualRow, namaCol).setValue(updateData.NAMA);
    }
    
    if (updateData.KELAS) {
      var kelasCol = headers.indexOf("KELAS") + 1;
      sheet.getRange(actualRow, kelasCol).setValue(updateData.KELAS);
    }
    
    return { success: true, message: "Profil berhasil diperbarui!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ==================== BUKU API ====================

function getBooks() {
  try {
    return { success: true, data: getSheetDataAsObjects("BUKU") };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function addBook(bookData) {
  try {
    var db = getDb();
    var sheet = db.getSheetByName("BUKU");
    var books = getSheetDataAsObjects("BUKU");
    
    var newId = "BKP-" + String(books.length + 1).padStart(3, '0');
    
    var coverUrl = "https://images.unsplash.com/photo-1543002588-bfa74002ed7e?auto=format&fit=crop&q=80&w=400"; // default cover
    if (bookData.COVER_BASE64) {
      coverUrl = uploadFileToDrive(bookData.COVER_BASE64, newId + "_cover", COVER_FOLDER_NAME, bookData.COVER_MIME || "image/png");
    }
    
    var pdfUrl = "";
    if (bookData.PDF_BASE64) {
      pdfUrl = uploadFileToDrive(bookData.PDF_BASE64, newId + "_" + bookData.JUDUL.replace(/[^a-z0-9]/gi, '_'), PDF_FOLDER_NAME, "application/pdf");
    } else {
      pdfUrl = bookData.LINK_PDF || "";
    }
    
    sheet.appendRow([
      newId,
      bookData.JUDUL,
      bookData.PENULIS,
      bookData.PENERBIT,
      bookData.TAHUN,
      bookData.KATEGORI,
      bookData.DESKRIPSI,
      coverUrl,
      pdfUrl,
      0, // JUMLAH_DIBACA
      bookData.STATUS || "Aktif",
      new Date().toISOString()
    ]);
    
    return { success: true, message: "Buku berhasil ditambahkan!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function editBook(bookId, bookData) {
  try {
    var db = getDb();
    var sheet = db.getSheetByName("BUKU");
    var books = getSheetDataAsObjects("BUKU");
    var rowIdx = books.findIndex(function(b) { return b.ID_BUKU === bookId; });
    
    if (rowIdx === -1) {
      return { success: false, message: "Buku tidak ditemukan!" };
    }
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var actualRow = rowIdx + 2;
    
    var fieldsToUpdate = ["JUDUL", "PENULIS", "PENERBIT", "TAHUN", "KATEGORI", "DESKRIPSI", "STATUS"];
    fieldsToUpdate.forEach(function(field) {
      if (bookData[field] !== undefined) {
        var col = headers.indexOf(field) + 1;
        sheet.getRange(actualRow, col).setValue(bookData[field]);
      }
    });
    
    if (bookData.COVER_BASE64) {
      var coverUrl = uploadFileToDrive(bookData.COVER_BASE64, bookId + "_cover", COVER_FOLDER_NAME, bookData.COVER_MIME || "image/png");
      var coverCol = headers.indexOf("COVER") + 1;
      sheet.getRange(actualRow, coverCol).setValue(coverUrl);
    }
    
    if (bookData.PDF_BASE64) {
      var pdfUrl = uploadFileToDrive(bookData.PDF_BASE64, bookId + "_" + bookData.JUDUL.replace(/[^a-z0-9]/gi, '_'), PDF_FOLDER_NAME, "application/pdf");
      var pdfCol = headers.indexOf("LINK_PDF") + 1;
      sheet.getRange(actualRow, pdfCol).setValue(pdfUrl);
    } else if (bookData.LINK_PDF !== undefined) {
      var pdfCol = headers.indexOf("LINK_PDF") + 1;
      sheet.getRange(actualRow, pdfCol).setValue(bookData.LINK_PDF);
    }
    
    return { success: true, message: "Buku berhasil diperbarui!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteBook(bookId) {
  try {
    var db = getDb();
    var sheet = db.getSheetByName("BUKU");
    var books = getSheetDataAsObjects("BUKU");
    var rowIdx = books.findIndex(function(b) { return b.ID_BUKU === bookId; });
    
    if (rowIdx === -1) {
      return { success: false, message: "Buku tidak ditemukan!" };
    }
    
    sheet.deleteRow(rowIdx + 2);
    return { success: true, message: "Buku berhasil dihapus!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function incrementReadCount(bookId) {
  try {
    var db = getDb();
    var sheet = db.getSheetByName("BUKU");
    var books = getSheetDataAsObjects("BUKU");
    var rowIdx = books.findIndex(function(b) { return b.ID_BUKU === bookId; });
    
    if (rowIdx === -1) return { success: false };
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var readCol = headers.indexOf("JUMLAH_DIBACA") + 1;
    var currentCount = Number(books[rowIdx].JUMLAH_DIBACA || 0);
    
    sheet.getRange(rowIdx + 2, readCol).setValue(currentCount + 1);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ==================== KATEGORI API ====================

function getCategories() {
  try {
    return { success: true, data: getSheetDataAsObjects("KATEGORI") };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function addCategory(catData) {
  try {
    var db = getDb();
    var sheet = db.getSheetByName("KATEGORI");
    var cats = getSheetDataAsObjects("KATEGORI");
    
    var newId = "CAT-" + String(cats.length + 1).padStart(3, '0');
    
    sheet.appendRow([
      newId,
      catData.NAMA_KATEGORI,
      catData.ICON || "fa-bookmark",
      catData.WARNA || "blue"
    ]);
    
    return { success: true, message: "Kategori berhasil ditambahkan!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function editCategory(catId, catData) {
  try {
    var db = getDb();
    var sheet = db.getSheetByName("KATEGORI");
    var cats = getSheetDataAsObjects("KATEGORI");
    var rowIdx = cats.findIndex(function(c) { return c.ID_KATEGORI === catId; });
    
    if (rowIdx === -1) return { success: false, message: "Kategori tidak ditemukan!" };
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var actualRow = rowIdx + 2;
    
    if (catData.NAMA_KATEGORI) {
      sheet.getRange(actualRow, headers.indexOf("NAMA_KATEGORI") + 1).setValue(catData.NAMA_KATEGORI);
    }
    if (catData.ICON) {
      sheet.getRange(actualRow, headers.indexOf("ICON") + 1).setValue(catData.ICON);
    }
    if (catData.WARNA) {
      sheet.getRange(actualRow, headers.indexOf("WARNA") + 1).setValue(catData.WARNA);
    }
    
    return { success: true, message: "Kategori berhasil diperbarui!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteCategory(catId) {
  try {
    var db = getDb();
    var sheet = db.getSheetByName("KATEGORI");
    var cats = getSheetDataAsObjects("KATEGORI");
    var rowIdx = cats.findIndex(function(c) { return c.ID_KATEGORI === catId; });
    
    if (rowIdx === -1) return { success: false, message: "Kategori tidak ditemukan!" };
    
    sheet.deleteRow(rowIdx + 2);
    return { success: true, message: "Kategori berhasil dihapus!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ==================== ANGGOTA API ====================

function getMembers() {
  try {
    var users = getSheetDataAsObjects("USERS");
    // Filter to return only users with role "Pembaca" or display all for management, but exclude actual passwords
    var members = users.map(function(u) {
      var copy = Object.assign({}, u);
      delete copy.PASSWORD;
      return copy;
    });
    return { success: true, data: members };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function addMember(memberData) {
  try {
    var db = getDb();
    var sheet = db.getSheetByName("USERS");
    var users = getSheetDataAsObjects("USERS");
    
    var dup = users.some(function(u) {
      return u.USERNAME.toLowerCase() === memberData.USERNAME.toLowerCase();
    });
    if (dup) {
      return { success: false, message: "Username sudah digunakan!" };
    }
    
    var newId = "USR-" + String(users.length + 1).padStart(3, '0');
    
    var photoUrl = "";
    if (memberData.PHOTO_BASE64) {
      photoUrl = uploadFileToDrive(memberData.PHOTO_BASE64, newId + "_avatar", AVATAR_FOLDER_NAME, memberData.PHOTO_MIME || "image/png");
    }
    
    var salt = "pustaka_premium_salt";
    var hashed = hashPassword(memberData.PASSWORD || "123456", salt); // default password
    
    sheet.appendRow([
      newId,
      memberData.NAMA,
      memberData.USERNAME,
      hashed,
      memberData.ROLE || "Pembaca",
      memberData.KELAS || "Umum",
      photoUrl,
      memberData.STATUS || "Aktif",
      new Date().toISOString()
    ]);
    
    return { success: true, message: "Anggota berhasil ditambahkan!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function editMember(memberId, memberData) {
  try {
    var db = getDb();
    var sheet = db.getSheetByName("USERS");
    var users = getSheetDataAsObjects("USERS");
    var rowIdx = users.findIndex(function(u) { return u.ID_USER === memberId; });
    
    if (rowIdx === -1) return { success: false, message: "Anggota tidak ditemukan!" };
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var actualRow = rowIdx + 2;
    
    var fields = ["NAMA", "ROLE", "KELAS", "STATUS"];
    fields.forEach(function(field) {
      if (memberData[field] !== undefined) {
        sheet.getRange(actualRow, headers.indexOf(field) + 1).setValue(memberData[field]);
      }
    });
    
    if (memberData.PHOTO_BASE64) {
      var photoUrl = uploadFileToDrive(memberData.PHOTO_BASE64, memberId + "_avatar", AVATAR_FOLDER_NAME, memberData.PHOTO_MIME || "image/png");
      sheet.getRange(actualRow, headers.indexOf("FOTO") + 1).setValue(photoUrl);
    }
    
    if (memberData.PASSWORD) {
      var hashed = hashPassword(memberData.PASSWORD, "pustaka_premium_salt");
      sheet.getRange(actualRow, headers.indexOf("PASSWORD") + 1).setValue(hashed);
    }
    
    return { success: true, message: "Anggota berhasil diperbarui!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteMember(memberId) {
  try {
    var db = getDb();
    var sheet = db.getSheetByName("USERS");
    var users = getSheetDataAsObjects("USERS");
    var rowIdx = users.findIndex(function(u) { return u.ID_USER === memberId; });
    
    if (rowIdx === -1) return { success: false, message: "Anggota tidak ditemukan!" };
    
    sheet.deleteRow(rowIdx + 2);
    return { success: true, message: "Anggota berhasil dihapus!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function resetMemberPassword(memberId) {
  return editMember(memberId, { PASSWORD: "pembaca123" });
}

// ==================== RIWAYAT BACA & STATS API ====================

function getHistory() {
  try {
    return { success: true, data: getSheetDataAsObjects("RIWAYAT_BACA") };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function addHistory(username, bookId, bookTitle) {
  try {
    var db = getDb();
    var sheet = db.getSheetByName("RIWAYAT_BACA");
    var lastRow = sheet.getLastRow();
    
    var newId = "RWY-" + String(lastRow).padStart(4, '0');
    
    sheet.appendRow([
      newId,
      username,
      bookId,
      bookTitle,
      new Date().toISOString()
    ]);
    
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ==================== BACKUP, RESTORE & SETTINGS API ====================

function getAppSettings() {
  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    var settings = scriptProperties.getProperties();
    return {
      success: true,
      data: {
        libraryName: settings.libraryName || "Pustaka Digital Premium",
        themeColor: settings.themeColor || "blue",
        logoUrl: settings.logoUrl || "https://images.unsplash.com/photo-1516979187457-637abb4f9353?auto=format&fit=crop&q=80&w=100",
        allowDownload: settings.allowDownload !== "false"
      }
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function saveAppSettings(settings) {
  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperties({
      libraryName: settings.libraryName || "Pustaka Digital Premium",
      themeColor: settings.themeColor || "blue",
      logoUrl: settings.logoUrl || "",
      allowDownload: String(settings.allowDownload)
    });
    return { success: true, message: "Pengaturan berhasil disimpan!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function backupDatabase() {
  try {
    var data = {
      USERS: getSheetDataAsObjects("USERS"),
      BUKU: getSheetDataAsObjects("BUKU"),
      KATEGORI: getSheetDataAsObjects("KATEGORI"),
      RIWAYAT_BACA: getSheetDataAsObjects("RIWAYAT_BACA"),
      backupTime: new Date().toISOString()
    };
    
    var jsonStr = JSON.stringify(data, null, 2);
    var folder = getFolder("Pustaka_Backups");
    var file = folder.createFile("Backup_DB_" + new Date().toISOString().slice(0,10) + ".json", jsonStr, "application/json");
    
    return { 
      success: true, 
      message: "Backup berhasil dibuat!", 
      backupUrl: file.getUrl(),
      backupData: jsonStr // return data string directly so users can also download local files
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function restoreDatabase(backupJsonStr) {
  try {
    var data = JSON.parse(backupJsonStr);
    var db = getDb();
    
    var sheets = ["USERS", "BUKU", "KATEGORI", "RIWAYAT_BACA"];
    sheets.forEach(function(sheetName) {
      if (!data[sheetName]) return;
      
      var sheet = db.getSheetByName(sheetName);
      if (sheet) {
        // Clear all except headers
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          sheet.deleteRows(2, lastRow - 1);
        }
      } else {
        sheet = db.insertSheet(sheetName);
      }
      
      var rows = data[sheetName];
      if (rows.length === 0) return;
      
      // Get header columns in order
      var headers = Object.keys(rows[0]);
      
      // Reset headers in sheet just in case
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      rows.forEach(function(row) {
        var rowValues = headers.map(function(h) { return row[h]; });
        sheet.appendRow(rowValues);
      });
    });
    
    return { success: true, message: "Database berhasil direstore dari backup!" };
  } catch (err) {
    return { success: false, message: "Restore gagal: " + err.message };
  }
}
