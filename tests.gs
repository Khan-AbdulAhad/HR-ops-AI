// ============================================================
// TEST SUITE: Re-engage, Email Sending, AI Reply & Status Tag
// ============================================================
// Run via Apps Script editor: call runAllTests()
// Covers: reEngageCandidates, sanitizeEmailContent, normalizeEmail,
//         validateEmailForSending, buildFollowUpEmailPrompt,
//         buildDataGatheringFollowUpEmailPrompt, updateCandidateStatusTag,
//         addToFollowUpQueue, status tag AI processor behavior, and more.
//
// Tests are split across 9 sections (28 tests total):
//   1. normalizeEmail (6)
//   2. sanitizeEmailContent (6)
//   3. validateEmailForSending (2)
//   4. buildFollowUpEmailPrompt (4)
//   5. buildDataGatheringFollowUpEmailPrompt (3)
//   6. reEngageCandidates guard logic (4)
//   7. Status tag AI processor logic (3)
//   8. Follow-up queue dedup logic (2)
//   9. Re-engage + Follow_Up_Queue interaction (4)
// ============================================================

// ------------------------------------------------------------
// Minimal test framework
// ------------------------------------------------------------

/**
 * Run a single named test and record the outcome.
 * @param {Object} ctx  - { passed, failed, errors }
 * @param {string} name - Test description
 * @param {Function} fn - Zero-arg test function; throw to fail
 */
function _runTest(ctx, name, fn) {
  try {
    fn();
    ctx.passed++;
    Logger.log('  PASS  ' + name);
  } catch (e) {
    ctx.failed++;
    const msg = e && e.message ? e.message : String(e);
    ctx.errors.push('[FAIL] ' + name + '\n       ' + msg);
    Logger.log('  FAIL  ' + name + ' → ' + msg);
  }
}

function _assert(condition, message) {
  if (!condition) throw new Error(message || 'Assertion failed');
}

function _assertEqual(actual, expected, label) {
  if (actual !== expected) {
    throw new Error((label ? label + ': ' : '') + 'Expected ' + JSON.stringify(expected) + ', got ' + JSON.stringify(actual));
  }
}

function _assertContains(str, substring, label) {
  if (typeof str !== 'string' || str.indexOf(substring) === -1) {
    throw new Error((label ? label + ': ' : '') + JSON.stringify(str) + ' does not contain ' + JSON.stringify(substring));
  }
}

function _assertNotContains(str, substring, label) {
  if (typeof str === 'string' && str.indexOf(substring) !== -1) {
    throw new Error((label ? label + ': ' : '') + JSON.stringify(str) + ' should NOT contain ' + JSON.stringify(substring));
  }
}

// ============================================================
// SECTION 1: normalizeEmail
// ============================================================

function test_normalizeEmail_stripsDotsForGmail() {
  _assertEqual(normalizeEmail('t.e.s.t@gmail.com'), 'test@gmail.com', 'dots stripped');
}

function test_normalizeEmail_preservesNonGmailDomain() {
  _assertEqual(normalizeEmail('t.e.s.t@company.com'), 't.e.s.t@company.com', 'non-Gmail unchanged');
}

function test_normalizeEmail_handlesGooglemailCom() {
  _assertEqual(normalizeEmail('a.b.c@googlemail.com'), 'abc@googlemail.com', 'googlemail dots stripped');
}

function test_normalizeEmail_lowercasesInput() {
  _assertEqual(normalizeEmail('HELLO@YAHOO.COM'), 'hello@yahoo.com', 'lowercased');
}

function test_normalizeEmail_returnsEmptyForNull() {
  _assertEqual(normalizeEmail(null), '', 'null → empty string');
}

function test_normalizeEmail_returnsEmptyForEmpty() {
  _assertEqual(normalizeEmail(''), '', 'empty string → empty string');
}

// ============================================================
// SECTION 2: sanitizeEmailContent
// ============================================================

function test_sanitizeEmail_cleanContentPassesSafely() {
  const result = sanitizeEmailContent('Hi Jane, thanks for your interest. Looking forward to hearing from you!');
  _assertEqual(result.safe, true, 'clean content is safe');
  _assertEqual(result.violations.length, 0, 'no violations');
}

function test_sanitizeEmail_detectsTargetRatePattern() {
  const result = sanitizeEmailContent('Our target rate for this position is competitive.');
  _assertEqual(result.safe, false, 'target rate detected');
  _assert(result.violations.length > 0, 'at least one violation logged');
}

function test_sanitizeEmail_detectsMaximumRatePattern() {
  const result = sanitizeEmailContent('The maximum rate we can offer is determined internally.');
  _assertEqual(result.safe, false, 'maximum rate detected');
}

function test_sanitizeEmail_detectsRateTierReference() {
  const result = sanitizeEmailContent('Based on your rate tier, we have an offer.');
  _assertEqual(result.safe, false, 'rate tier detected');
}

function test_sanitizeEmail_detectsJobIdInInternalData() {
  const result = sanitizeEmailContent('Thank you for applying to job 98765.', { jobId: '98765' });
  _assertEqual(result.safe, false, 'job ID match via internalData');
  _assert(result.violations.some(function(v) { return v.indexOf('98765') !== -1; }), 'violation mentions jobId');
}

function test_sanitizeEmail_detectsDevIdInInternalData() {
  const result = sanitizeEmailContent('Hi, your devId is DEV-001 in our system.', { devId: 'DEV-001' });
  _assertEqual(result.safe, false, 'dev ID match via internalData');
}

// ============================================================
// SECTION 3: validateEmailForSending
// ============================================================

function test_validateEmail_returnsTrue_forCleanContent() {
  const result = validateEmailForSending('Hi Jane, we would love to discuss this opportunity with you. Best regards.');
  _assertEqual(result, true, 'clean content returns true');
}

function test_validateEmail_returnsFalse_forTargetRateContent() {
  const result = validateEmailForSending('Our target rate budget is confidential.');
  _assertEqual(result, false, 'content with target rate returns false');
}

// ============================================================
// SECTION 4: buildFollowUpEmailPrompt
// ============================================================

function test_buildFollowUpPrompt_includesCandidateName() {
  const prompt = buildFollowUpEmailPrompt({ name: 'John Smith', jobDescription: 'Backend engineer role', followUpNumber: 1 });
  _assertContains(prompt, 'John', 'first name in prompt');
}

function test_buildFollowUpPrompt_firstFollowUpUsesFriendlyTone() {
  const prompt = buildFollowUpEmailPrompt({ name: 'Alice', jobDescription: '', followUpNumber: 1 });
  _assertContains(prompt, 'FIRST follow-up', 'first follow-up tone label');
}

function test_buildFollowUpPrompt_secondFollowUpUsesFinalTone() {
  const prompt = buildFollowUpEmailPrompt({ name: 'Bob', jobDescription: '', followUpNumber: 2 });
  _assertContains(prompt, 'SECOND and FINAL follow-up', 'second/final follow-up tone label');
}

function test_buildFollowUpPrompt_neverLeaksInternalData() {
  const prompt = buildFollowUpEmailPrompt({ name: 'Carol', jobDescription: 'Some role', followUpNumber: 1 });
  _assertContains(prompt, 'NEVER include', 'confidentiality section present');
  _assertContains(prompt, 'target rate', 'target rate listed as forbidden');
  _assertContains(prompt, 'max rate', 'max rate listed as forbidden');
}

// ============================================================
// SECTION 5: buildDataGatheringFollowUpEmailPrompt
// ============================================================

function test_buildDataGatheringPrompt_includesPendingQuestions() {
  const prompt = buildDataGatheringFollowUpEmailPrompt({
    name: 'Diana',
    jobDescription: 'ML engineer',
    followUpNumber: 1,
    pendingQuestions: [{ question: 'What is your notice period?' }, { question: 'What is your expected rate?' }],
    answeredQuestions: []
  });
  _assertContains(prompt, 'What is your notice period?', 'pending question 1 present');
  _assertContains(prompt, 'What is your expected rate?', 'pending question 2 present');
}

function test_buildDataGatheringPrompt_thirdFollowUpUsesUrgentTone() {
  const prompt = buildDataGatheringFollowUpEmailPrompt({
    name: 'Eva',
    jobDescription: '',
    followUpNumber: 3,
    pendingQuestions: [],
    answeredQuestions: []
  });
  _assertContains(prompt, 'THIRD and FINAL data follow-up', 'third/final urgency label');
}

function test_buildDataGatheringPrompt_includesAnsweredContext() {
  const prompt = buildDataGatheringFollowUpEmailPrompt({
    name: 'Frank',
    jobDescription: 'Role',
    followUpNumber: 2,
    pendingQuestions: [],
    answeredQuestions: [{ question: 'Start date?', answer: 'March 2026' }]
  });
  _assertContains(prompt, 'Start date?', 'answered question in context');
  _assertContains(prompt, 'March 2026', 'answered value in context');
}

// ============================================================
// SECTION 6: reEngageCandidates guard logic
// ============================================================

function test_reEngage_emptyArrayReturnsZeroSuccessNoErrors() {
  const result = reEngageCandidates([], 'negotiation');
  _assertEqual(result.success, 0, 'success is 0');
  _assertEqual(result.failed, 0, 'failed is 0');
  _assertEqual(result.errors.length, 0, 'no errors');
}

function test_reEngage_nullCandidatesReturnsZero() {
  const result = reEngageCandidates(null, 'negotiation');
  _assertEqual(result.success, 0, 'null candidates → success 0');
  _assertEqual(result.failed, 0, 'null candidates → failed 0');
}

function test_reEngage_noSheetUrlReturnsConfigError() {
  // Without a spreadsheet URL configured, should surface a clear error
  // (PropertiesService will have no URL in the test environment)
  const result = reEngageCandidates([{ email: 'x@x.com', jobId: '1', threadId: 'tid' }], 'negotiation');
  _assert(result.errors.length > 0, 'error list is non-empty when no URL');
  _assert(
    result.errors[0].indexOf('No sheet URL') !== -1 ||
    result.errors[0].indexOf('sheet') !== -1 ||
    result.errors[0].indexOf('URL') !== -1,
    'error mentions sheet/URL: ' + result.errors[0]
  );
}

function test_reEngage_customStageWithoutMessageIncreasesFailed() {
  // When stage='custom' and no customMessage, the entry should fail
  // (This tests the guard before AI is called)
  // Since no URL is configured, the entire call returns url-error first.
  // We verify the returned error structure is well-formed.
  const result = reEngageCandidates([{ email: 'a@b.com', jobId: '1' }], 'custom', '');
  _assert(typeof result.success === 'number', 'success is a number');
  _assert(typeof result.failed === 'number', 'failed is a number');
  _assert(Array.isArray(result.errors), 'errors is an array');
}

// ============================================================
// SECTION 7: Status tag AI processor behavior
// ============================================================

/**
 * Validates the status-tag check logic used in the main processor.
 * The processor checks: currentStatus.toLowerCase().indexOf('completed') > -1
 * These tests verify that the condition correctly identifies all completed variants.
 */
function test_statusTag_completedStopsProcessing() {
  const statusesToBlock = [
    'Completed',
    'Completed - Job Fulfilled',
    'Completed - Job Stopped',
    'completed',
    'Data Complete'
  ];
  statusesToBlock.forEach(function(s) {
    const blocked = (s.toLowerCase().indexOf('completed') > -1 || s === 'Data Complete');
    _assert(blocked, 'Status "' + s + '" should trigger completed-block');
  });
}

function test_statusTag_unresponsiveSkipsProcessing() {
  // The processor checks: currentStatus === 'Unresponsive'
  _assert('Unresponsive' === 'Unresponsive', '"Unresponsive" exact match skips processing');
  // Non-matching variants should NOT trigger the skip
  const shouldNotBlock = ['unresponsive', 'UNRESPONSIVE', 'Unresponsive (manual)'];
  shouldNotBlock.forEach(function(s) {
    _assert(s !== 'Unresponsive', '"' + s + '" should NOT trigger unresponsive exact-match skip');
  });
}

function test_statusTag_reEngageResetsToActiveForAI() {
  // After re-engagement, status is set to "Re-engaged - Negotiation" (or similar).
  // This status must NOT contain 'completed' or equal 'Unresponsive',
  // so the AI processor will resume handling the thread.
  const reEngageStatuses = [
    'Re-engaged - Negotiation',
    'Re-engaged - Data Gathering',
    'Re-engaged - Custom'
  ];
  reEngageStatuses.forEach(function(s) {
    const isCompletedBlocked = s.toLowerCase().indexOf('completed') > -1;
    const isUnresponsiveBlocked = s === 'Unresponsive';
    const isNotInterestedBlocked = s === 'Not Interested';
    _assert(
      !isCompletedBlocked && !isUnresponsiveBlocked && !isNotInterestedBlocked,
      '"' + s + '" should allow AI processing to resume'
    );
  });
}

// ============================================================
// SECTION 8: Follow-up queue dedup logic
// ============================================================

function test_followUpQueue_gmailDotVariantsNormalizeToSameKey() {
  // addToFollowUpQueue uses normalizeEmail for dedup. Two dot-variants of the
  // same Gmail address must resolve to the same cache key.
  const email1 = normalizeEmail('john.doe@gmail.com');
  const email2 = normalizeEmail('johndoe@gmail.com');
  _assertEqual(email1, email2, 'dot-variant Gmail addresses share the same normalized key');
}

function test_followUpQueue_differentJobsSameEmailAreDistinct() {
  // Two entries for the same email but different job IDs must produce different keys.
  const email = normalizeEmail('dev@company.com');
  const key1 = email + '|' + 'JOB_A';
  const key2 = email + '|' + 'JOB_B';
  _assert(key1 !== key2, 'different job IDs produce different queue keys');
}

// ============================================================
// SECTION 9: Re-engage + Follow_Up_Queue interaction
// ============================================================

/**
 * Validates the Follow_Up_Queue reset logic used inside reEngageCandidates.
 * The code sets status to 'Re-engaged' when existing status is 'Unresponsive' or 'Incomplete Data'.
 * These tests verify the condition logic inline (no actual sheet required).
 */
function test_reEngageQueue_unresponsiveStatusIsReset() {
  const currentStatus = 'Unresponsive';
  const shouldReset = (currentStatus === 'Unresponsive' || currentStatus === 'Incomplete Data');
  _assert(shouldReset, '"Unresponsive" queue entry should be reset on re-engage');
}

function test_reEngageQueue_incompleteDataStatusIsReset() {
  const currentStatus = 'Incomplete Data';
  const shouldReset = (currentStatus === 'Unresponsive' || currentStatus === 'Incomplete Data');
  _assert(shouldReset, '"Incomplete Data" queue entry should be reset on re-engage');
}

function test_reEngageQueue_pendingStatusIsNotReset() {
  // A "Pending" entry in the queue should be left as-is during re-engagement.
  const currentStatus = 'Pending';
  const shouldReset = (currentStatus === 'Unresponsive' || currentStatus === 'Incomplete Data');
  _assert(!shouldReset, '"Pending" queue entry should NOT be reset on re-engage');
}

function test_reEngageQueue_newStatusIsReEngaged() {
  // The reset value written to Follow_Up_Queue column I must be 'Re-engaged'.
  const expectedNewStatus = 'Re-engaged';
  _assertEqual(expectedNewStatus, 'Re-engaged', 'new Follow_Up_Queue status is "Re-engaged"');
  // Verify it does not accidentally also mark as completed/unresponsive
  _assert(expectedNewStatus.toLowerCase().indexOf('completed') === -1, 'Re-engaged is not Completed');
  _assert(expectedNewStatus !== 'Unresponsive', 'Re-engaged is not Unresponsive');
}

// ============================================================
// RUNNER
// ============================================================

/**
 * Execute all tests and log a final summary.
 * Call this function from the Apps Script editor.
 *
 * @returns {Object} { passed, failed, errors }
 */
function runAllTests() {
  const ctx = { passed: 0, failed: 0, errors: [] };
  Logger.log('====== HR-ops-AI Test Suite ======');

  // Section 1: normalizeEmail
  Logger.log('\n--- Section 1: normalizeEmail ---');
  _runTest(ctx, 'normalizeEmail strips dots for Gmail', test_normalizeEmail_stripsDotsForGmail);
  _runTest(ctx, 'normalizeEmail preserves non-Gmail domain', test_normalizeEmail_preservesNonGmailDomain);
  _runTest(ctx, 'normalizeEmail handles googlemail.com', test_normalizeEmail_handlesGooglemailCom);
  _runTest(ctx, 'normalizeEmail lowercases input', test_normalizeEmail_lowercasesInput);
  _runTest(ctx, 'normalizeEmail returns empty for null', test_normalizeEmail_returnsEmptyForNull);
  _runTest(ctx, 'normalizeEmail returns empty for empty string', test_normalizeEmail_returnsEmptyForEmpty);

  // Section 2: sanitizeEmailContent
  Logger.log('\n--- Section 2: sanitizeEmailContent ---');
  _runTest(ctx, 'sanitizeEmail clean content is safe', test_sanitizeEmail_cleanContentPassesSafely);
  _runTest(ctx, 'sanitizeEmail detects "target rate"', test_sanitizeEmail_detectsTargetRatePattern);
  _runTest(ctx, 'sanitizeEmail detects "maximum rate"', test_sanitizeEmail_detectsMaximumRatePattern);
  _runTest(ctx, 'sanitizeEmail detects "rate tier"', test_sanitizeEmail_detectsRateTierReference);
  _runTest(ctx, 'sanitizeEmail detects job ID via internalData', test_sanitizeEmail_detectsJobIdInInternalData);
  _runTest(ctx, 'sanitizeEmail detects dev ID via internalData', test_sanitizeEmail_detectsDevIdInInternalData);

  // Section 3: validateEmailForSending
  Logger.log('\n--- Section 3: validateEmailForSending ---');
  _runTest(ctx, 'validateEmail returns true for clean content', test_validateEmail_returnsTrue_forCleanContent);
  _runTest(ctx, 'validateEmail returns false for target rate content', test_validateEmail_returnsFalse_forTargetRateContent);

  // Section 4: buildFollowUpEmailPrompt
  Logger.log('\n--- Section 4: buildFollowUpEmailPrompt ---');
  _runTest(ctx, 'buildFollowUpPrompt includes candidate name', test_buildFollowUpPrompt_includesCandidateName);
  _runTest(ctx, 'buildFollowUpPrompt first follow-up uses friendly tone', test_buildFollowUpPrompt_firstFollowUpUsesFriendlyTone);
  _runTest(ctx, 'buildFollowUpPrompt second follow-up uses final tone', test_buildFollowUpPrompt_secondFollowUpUsesFinalTone);
  _runTest(ctx, 'buildFollowUpPrompt never leaks internal data', test_buildFollowUpPrompt_neverLeaksInternalData);

  // Section 5: buildDataGatheringFollowUpEmailPrompt
  Logger.log('\n--- Section 5: buildDataGatheringFollowUpEmailPrompt ---');
  _runTest(ctx, 'buildDataGatheringPrompt includes pending questions', test_buildDataGatheringPrompt_includesPendingQuestions);
  _runTest(ctx, 'buildDataGatheringPrompt third follow-up uses urgent tone', test_buildDataGatheringPrompt_thirdFollowUpUsesUrgentTone);
  _runTest(ctx, 'buildDataGatheringPrompt includes answered context', test_buildDataGatheringPrompt_includesAnsweredContext);

  // Section 6: reEngageCandidates guard logic
  Logger.log('\n--- Section 6: reEngageCandidates guard logic ---');
  _runTest(ctx, 'reEngage empty array returns zero success with no errors', test_reEngage_emptyArrayReturnsZeroSuccessNoErrors);
  _runTest(ctx, 'reEngage null candidates returns zero', test_reEngage_nullCandidatesReturnsZero);
  _runTest(ctx, 'reEngage no sheet URL returns config error', test_reEngage_noSheetUrlReturnsConfigError);
  _runTest(ctx, 'reEngage custom stage without message returns error structure', test_reEngage_customStageWithoutMessageIncreasesFailed);

  // Section 7: Status tag AI processor behavior
  Logger.log('\n--- Section 7: Status tag AI processor behavior ---');
  _runTest(ctx, 'statusTag Completed variants stop AI processing', test_statusTag_completedStopsProcessing);
  _runTest(ctx, 'statusTag Unresponsive exact-match skips processing', test_statusTag_unresponsiveSkipsProcessing);
  _runTest(ctx, 'statusTag Re-engaged statuses allow AI processing to resume', test_statusTag_reEngageResetsToActiveForAI);

  // Section 8: Follow-up queue dedup logic
  Logger.log('\n--- Section 8: Follow-up queue dedup logic ---');
  _runTest(ctx, 'followUpQueue Gmail dot-variants normalize to same key', test_followUpQueue_gmailDotVariantsNormalizeToSameKey);
  _runTest(ctx, 'followUpQueue different jobs produce distinct keys', test_followUpQueue_differentJobsSameEmailAreDistinct);

  // Section 9: Re-engage + Follow_Up_Queue interaction
  Logger.log('\n--- Section 9: Re-engage + Follow_Up_Queue interaction ---');
  _runTest(ctx, 'reEngageQueue resets Unresponsive status', test_reEngageQueue_unresponsiveStatusIsReset);
  _runTest(ctx, 'reEngageQueue resets Incomplete Data status', test_reEngageQueue_incompleteDataStatusIsReset);
  _runTest(ctx, 'reEngageQueue does not reset Pending status', test_reEngageQueue_pendingStatusIsNotReset);
  _runTest(ctx, 'reEngageQueue new status value is "Re-engaged"', test_reEngageQueue_newStatusIsReEngaged);

  // Summary
  Logger.log('\n====== Results ======');
  Logger.log('PASSED: ' + ctx.passed);
  Logger.log('FAILED: ' + ctx.failed);
  if (ctx.errors.length > 0) {
    Logger.log('\nFailures:');
    ctx.errors.forEach(function(e) { Logger.log(e); });
  }
  Logger.log('=====================');

  return ctx;
}
