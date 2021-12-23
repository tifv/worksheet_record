var ActionLock = function() { // begin namespace

function acquire() {
  var lock = LockService.getDocumentLock();
  var success = lock.tryLock(300);
  if (success)
    return lock;
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Ожидание завершения других операций." );
  var success = lock.tryLock(5000);
  if (success)
    return lock;
  throw new ReportError(
    "Не удалось получить доступ к таблице. " +
    "Другие операции (возможно, запущенные другими редакторами) " +
    "создают опасность одновременного редактирования." );
}

function with_lock(operator) {
  // var lock = acquire();
  try {
    return operator();
  } finally {
    // lock.releaseLock();
  }
}

return {acquire, with_lock};
}();

