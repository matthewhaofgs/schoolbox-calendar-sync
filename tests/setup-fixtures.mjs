export async function completeStoredSetup(storage, options = {}) {
  const actor = options.actor ?? "test:setup";
  const schoolbox = await storage.getStoredSchoolboxConnection();
  await storage.recordConnectionVerified("schoolbox", actor, schoolbox.credentialVersion);

  if (options.google !== false) {
    const google = await storage.getStoredGoogleConnection();
    await storage.recordConnectionVerified("google", actor, google.credentialVersion);
  }

  if (options.microsoft) {
    const microsoft = await storage.getStoredMicrosoftConnection();
    await storage.recordMicrosoftAdminConsent(actor, microsoft);
  }

  return storage.saveConfig({
    schoolboxSetupCompleted: true,
    ...(options.google !== false ? { googleSetupCompleted: true, googleEnabled: true } : {}),
    ...(options.microsoft ? { microsoftSetupCompleted: true, microsoftEnabled: true } : {}),
  }, actor);
}
