/**
 * Negotiation Toggle Tests
 *
 * Tests to verify that when negotiation is disabled, the AI does NOT negotiate
 * even when follow-up or data gathering (or both) are enabled.
 *
 * Also includes 5 edge case scenarios that could go wrong.
 */

// Test utilities
function createMockTestData(overrides = {}) {
  return {
    type: 'multifunction',
    devName: 'John Smith',
    devEmail: 'john.smith@example.com',
    devCountry: 'United States',
    jobId: '51000',
    jobDesc: 'Senior Software Engineer - AI/ML focus',
    candidateReply: '',
    targetRate: 50,
    maxRate: 60,
    attempt: 1,
    followUpNumber: 1,
    pendingQuestions: 'LinkedIn URL, Expected availability',
    functions: {
      negotiation: true,
      followup: true,
      datagathering: true
    },
    conversationHistory: null,
    conversationContext: null,
    ...overrides
  };
}

function analyzeResponse(response) {
  const negotiationIndicators = [
    /\$\d+\s*\/?\s*(hr|hour)/i,          // Rate mentions like "$50/hr"
    /rate/i,                              // Word "rate"
    /hourly/i,                            // "hourly"
    /compensation/i,                      // "compensation"
    /we can offer/i,                      // Offering rate
    /budget/i,                            // Budget mentions
    /salary/i,                            // Salary
    /pay/i,                               // Pay
    /offer.*\$\d+/i,                      // "offer $X"
  ];

  const dataGatheringIndicators = [
    /linkedin/i,
    /profile/i,
    /availability/i,
    /when.*start/i,
    /share.*details/i,
    /provide.*information/i,
  ];

  const followUpIndicators = [
    /follow(ing)?\s*up/i,
    /checking in/i,
    /haven't heard/i,
    /just wanted to/i,
    /reaching out again/i,
  ];

  return {
    containsNegotiation: negotiationIndicators.some(pattern => pattern.test(response)),
    containsDataGathering: dataGatheringIndicators.some(pattern => pattern.test(response)),
    containsFollowUp: followUpIndicators.some(pattern => pattern.test(response)),
    negotiationPatternMatches: negotiationIndicators.filter(p => p.test(response)).map(p => p.toString()),
  };
}

// ============================================================
// MAIN TEST SCENARIOS
// ============================================================

/**
 * TEST 1: Negotiation OFF + Follow-up ONLY
 * Expected: AI should NOT mention rates or negotiate
 */
const TEST_1_NEGOTIATION_OFF_FOLLOWUP_ONLY = {
  name: 'Negotiation OFF + Follow-up ON only',
  description: 'When negotiation is disabled but follow-up is enabled, AI should send follow-up WITHOUT any rate negotiation',
  testData: createMockTestData({
    candidateReply: 'I am interested in this role and looking for $55/hr',
    functions: {
      negotiation: false,  // <-- DISABLED
      followup: true,      // <-- ENABLED
      datagathering: false
    }
  }),
  expectedBehavior: {
    shouldNegotiate: false,
    shouldMentionRate: false,
    shouldSendFollowUp: true,
  },
  validate: (response, activeTypes) => {
    const analysis = analyzeResponse(response);
    const errors = [];

    // Should NOT contain negotiation content
    if (analysis.containsNegotiation) {
      errors.push(`FAIL: Response contains negotiation content when negotiation is OFF. Patterns: ${analysis.negotiationPatternMatches.join(', ')}`);
    }

    // Should NOT include 'negotiation' in activeTypes
    if (activeTypes && activeTypes.includes('negotiation')) {
      errors.push(`FAIL: 'negotiation' found in activeTypes when it should be disabled`);
    }

    // Should include 'followup' in activeTypes
    if (activeTypes && !activeTypes.includes('followup')) {
      errors.push(`WARNING: 'followup' not found in activeTypes`);
    }

    return {
      passed: errors.filter(e => e.startsWith('FAIL')).length === 0,
      errors
    };
  }
};

/**
 * TEST 2: Negotiation OFF + Data Gathering ONLY
 * Expected: AI should gather data but NOT negotiate rates
 */
const TEST_2_NEGOTIATION_OFF_DATAGATHERING_ONLY = {
  name: 'Negotiation OFF + Data Gathering ON only',
  description: 'When negotiation is disabled but data gathering is enabled, AI should collect data WITHOUT negotiating rates',
  testData: createMockTestData({
    candidateReply: 'My rate expectation is $55/hr and here is my LinkedIn: linkedin.com/in/johnsmith',
    pendingQuestions: 'Expected availability, Resume/CV, Phone number',
    functions: {
      negotiation: false,  // <-- DISABLED
      followup: false,
      datagathering: true  // <-- ENABLED
    }
  }),
  expectedBehavior: {
    shouldNegotiate: false,
    shouldMentionRate: false,
    shouldGatherData: true,
  },
  validate: (response, activeTypes) => {
    const analysis = analyzeResponse(response);
    const errors = [];

    // Should NOT contain negotiation content
    if (analysis.containsNegotiation) {
      errors.push(`FAIL: Response contains negotiation when negotiation is OFF. Patterns: ${analysis.negotiationPatternMatches.join(', ')}`);
    }

    // Should NOT include 'negotiation' in activeTypes
    if (activeTypes && activeTypes.includes('negotiation')) {
      errors.push(`FAIL: 'negotiation' found in activeTypes when it should be disabled`);
    }

    // Should include 'datagathering' in activeTypes
    if (activeTypes && !activeTypes.includes('datagathering')) {
      errors.push(`WARNING: 'datagathering' not found in activeTypes`);
    }

    // Should ask for missing information
    if (!analysis.containsDataGathering) {
      errors.push(`WARNING: Response doesn't seem to contain data gathering questions`);
    }

    return {
      passed: errors.filter(e => e.startsWith('FAIL')).length === 0,
      errors
    };
  }
};

/**
 * TEST 3: Negotiation OFF + Both Data Gathering & Follow-up ON
 * Expected: AI should gather data and follow up but NOT negotiate
 */
const TEST_3_NEGOTIATION_OFF_BOTH_ENABLED = {
  name: 'Negotiation OFF + Data Gathering ON + Follow-up ON',
  description: 'When negotiation is disabled but both other features are enabled, AI should NOT negotiate even if candidate mentions rates',
  testData: createMockTestData({
    candidateReply: 'I want $60/hr for this role. Is that possible?',
    pendingQuestions: 'LinkedIn URL, Start date availability',
    functions: {
      negotiation: false,  // <-- DISABLED
      followup: true,      // <-- ENABLED
      datagathering: true  // <-- ENABLED
    }
  }),
  expectedBehavior: {
    shouldNegotiate: false,
    shouldMentionRate: false,
    shouldGatherData: true,
  },
  validate: (response, activeTypes) => {
    const analysis = analyzeResponse(response);
    const errors = [];

    // Should NOT contain rate negotiation
    if (analysis.containsNegotiation) {
      errors.push(`FAIL: Response contains negotiation when negotiation is OFF. Patterns: ${analysis.negotiationPatternMatches.join(', ')}`);
    }

    // Should NOT include 'negotiation' in activeTypes
    if (activeTypes && activeTypes.includes('negotiation')) {
      errors.push(`FAIL: 'negotiation' found in activeTypes when disabled`);
    }

    return {
      passed: errors.filter(e => e.startsWith('FAIL')).length === 0,
      errors
    };
  }
};

// ============================================================
// EDGE CASE SCENARIOS - 5 Things That Could Go Wrong
// ============================================================

/**
 * EDGE CASE 1: Candidate explicitly asks about rates when negotiation is OFF
 * Risk: AI might "help" by discussing rates anyway
 */
const EDGE_1_CANDIDATE_ASKS_RATE = {
  name: 'Edge Case 1: Candidate directly asks about rate when negotiation OFF',
  description: 'Candidate asks "What is the hourly rate?" but negotiation is disabled - AI should NOT reveal rates',
  testData: createMockTestData({
    candidateReply: 'What is the hourly rate for this position? I need to know before proceeding.',
    pendingQuestions: 'LinkedIn URL',
    functions: {
      negotiation: false,  // <-- DISABLED
      followup: true,
      datagathering: true
    }
  }),
  expectedBehavior: {
    shouldNegotiate: false,
    shouldRevealRate: false,
    shouldRedirectOrDefer: true,
  },
  validate: (response, activeTypes) => {
    const analysis = analyzeResponse(response);
    const errors = [];

    // Should NOT contain specific rate offers like "$50/hr"
    const hasSpecificRate = /\$\d+\s*\/?\s*(hr|hour)/i.test(response);
    if (hasSpecificRate) {
      errors.push(`FAIL: Response contains specific rate offer when negotiation is OFF`);
    }

    // Should NOT include 'negotiation' in activeTypes
    if (activeTypes && activeTypes.includes('negotiation')) {
      errors.push(`FAIL: 'negotiation' found in activeTypes when disabled`);
    }

    return {
      passed: errors.filter(e => e.startsWith('FAIL')).length === 0,
      errors
    };
  }
};

/**
 * EDGE CASE 2: Conversation context has previous rate discussion but negotiation now OFF
 * Risk: AI might continue previous negotiation from context
 */
const EDGE_2_PREVIOUS_NEGOTIATION_CONTEXT = {
  name: 'Edge Case 2: Previous negotiation context but negotiation now OFF',
  description: 'Conversation history contains rate discussion, but negotiation is now disabled - AI should NOT continue negotiating',
  testData: createMockTestData({
    candidateReply: 'I thought about your offer of $45/hr and I can do $50/hr as my final offer.',
    pendingQuestions: 'Start date',
    functions: {
      negotiation: false,  // <-- DISABLED NOW
      followup: true,
      datagathering: true
    },
    conversationContext: {
      negotiationState: {
        attempt: 2,
        lastRate: 45,
        maxOffered: 50,
        rateAgreed: false
      },
      extractedData: {
        'LinkedIn URL': 'linkedin.com/in/test'
      }
    }
  }),
  expectedBehavior: {
    shouldContinueNegotiation: false,
    shouldAcceptRate: false,
    shouldOnlyGatherData: true,
  },
  validate: (response, activeTypes) => {
    const analysis = analyzeResponse(response);
    const errors = [];

    // Should NOT contain negotiation acceptance or counter
    const acceptancePatterns = /accept.*\$|agreed.*rate|we can proceed at|offer.*\$\d+/i;
    if (acceptancePatterns.test(response)) {
      errors.push(`FAIL: Response appears to continue negotiation when negotiation is OFF`);
    }

    if (activeTypes && activeTypes.includes('negotiation')) {
      errors.push(`FAIL: 'negotiation' found in activeTypes when it should be disabled`);
    }

    return {
      passed: errors.filter(e => e.startsWith('FAIL')).length === 0,
      errors
    };
  }
};

/**
 * EDGE CASE 3: Rate data in pending questions when negotiation OFF
 * Risk: Data gathering might inadvertently ask for rate expectations
 */
const EDGE_3_RATE_IN_PENDING_QUESTIONS = {
  name: 'Edge Case 3: Rate-related pending question with negotiation OFF',
  description: 'Pending questions include rate expectations but negotiation is OFF - should NOT ask about rates',
  testData: createMockTestData({
    candidateReply: 'I am very interested in the role.',
    pendingQuestions: 'LinkedIn URL, Rate expectations, Availability',
    functions: {
      negotiation: false,  // <-- DISABLED
      followup: false,
      datagathering: true
    }
  }),
  expectedBehavior: {
    shouldAskAboutRate: false,  // Rate questions should be skipped when negotiation OFF
    shouldAskOtherQuestions: true,
  },
  validate: (response, activeTypes) => {
    const errors = [];

    // Check if it asks about rate/compensation
    const asksAboutRate = /what.*rate|rate.*expect|hourly.*expect|compensation.*expect/i.test(response);
    if (asksAboutRate) {
      errors.push(`FAIL: Response asks about rate expectations when negotiation is OFF`);
    }

    if (activeTypes && activeTypes.includes('negotiation')) {
      errors.push(`FAIL: 'negotiation' found in activeTypes`);
    }

    return {
      passed: errors.filter(e => e.startsWith('FAIL')).length === 0,
      errors
    };
  }
};

/**
 * EDGE CASE 4: Candidate provides rate below max when negotiation OFF
 * Risk: AI might automatically "accept" the rate even though negotiation is disabled
 */
const EDGE_4_RATE_BELOW_MAX_NEGOTIATION_OFF = {
  name: 'Edge Case 4: Candidate gives rate below max but negotiation OFF',
  description: 'Candidate offers $40/hr (below $60 max) but negotiation is OFF - should NOT accept or discuss rate',
  testData: createMockTestData({
    candidateReply: 'I would be happy with $40/hr for this role.',
    targetRate: 50,
    maxRate: 60,
    pendingQuestions: 'LinkedIn URL, Start availability',
    functions: {
      negotiation: false,  // <-- DISABLED
      followup: false,
      datagathering: true
    }
  }),
  expectedBehavior: {
    shouldAcceptRate: false,
    shouldAcknowledgeRate: false,
    shouldOnlyGatherData: true,
  },
  validate: (response, activeTypes) => {
    const errors = [];

    // Should NOT accept or confirm the rate
    const acceptsRate = /accept.*\$40|confirmed.*\$40|great.*\$40|works.*\$40|noted.*rate|proceed.*\$40/i.test(response);
    if (acceptsRate) {
      errors.push(`FAIL: Response accepts/acknowledges rate when negotiation is OFF`);
    }

    if (activeTypes && activeTypes.includes('negotiation')) {
      errors.push(`FAIL: 'negotiation' found in activeTypes`);
    }

    return {
      passed: errors.filter(e => e.startsWith('FAIL')).length === 0,
      errors
    };
  }
};

/**
 * EDGE CASE 5: Toggle changed mid-conversation (negotiation was ON, now OFF)
 * Risk: System might not respect the new setting if there's cached state
 */
const EDGE_5_TOGGLE_CHANGED_MID_CONVERSATION = {
  name: 'Edge Case 5: Negotiation toggled OFF mid-conversation',
  description: 'Negotiation was previously ON (rate offered), now toggled OFF mid-conversation - should NOT continue negotiating',
  testData: createMockTestData({
    candidateReply: 'Your offer of $45/hr is too low. I need at least $55/hr.',
    targetRate: 50,
    maxRate: 60,
    pendingQuestions: 'LinkedIn URL',
    functions: {
      negotiation: false,  // <-- NOW DISABLED (was enabled before)
      followup: true,
      datagathering: true
    },
    conversationHistory: `
Previous AI message: "We can offer $45/hr for this position."
Candidate response: "That's too low, I need $55/hr."
`,
    conversationContext: {
      negotiationState: {
        attempt: 1,
        lastRate: 45,
        maxOffered: 60,
        rateAgreed: false
      }
    }
  }),
  expectedBehavior: {
    shouldCounterOffer: false,
    shouldAccept: false,
    shouldIgnoreRateDiscussion: true,
  },
  validate: (response, activeTypes) => {
    const errors = [];

    // Should NOT make counter offer
    const makesOffer = /offer.*\$\d+|we can.*\$\d+|counter.*\$|how about \$/i.test(response);
    if (makesOffer) {
      errors.push(`FAIL: Response makes rate offer when negotiation is now OFF`);
    }

    // Should NOT accept the rate
    const acceptsRate = /accept.*\$55|agreed.*\$55|works for us/i.test(response);
    if (acceptsRate) {
      errors.push(`FAIL: Response accepts rate when negotiation is OFF`);
    }

    if (activeTypes && activeTypes.includes('negotiation')) {
      errors.push(`FAIL: 'negotiation' found in activeTypes when it should be disabled`);
    }

    return {
      passed: errors.filter(e => e.startsWith('FAIL')).length === 0,
      errors
    };
  }
};

// ============================================================
// TEST RUNNER
// ============================================================

const ALL_TESTS = [
  // Main scenarios
  TEST_1_NEGOTIATION_OFF_FOLLOWUP_ONLY,
  TEST_2_NEGOTIATION_OFF_DATAGATHERING_ONLY,
  TEST_3_NEGOTIATION_OFF_BOTH_ENABLED,
  // Edge cases
  EDGE_1_CANDIDATE_ASKS_RATE,
  EDGE_2_PREVIOUS_NEGOTIATION_CONTEXT,
  EDGE_3_RATE_IN_PENDING_QUESTIONS,
  EDGE_4_RATE_BELOW_MAX_NEGOTIATION_OFF,
  EDGE_5_TOGGLE_CHANGED_MID_CONVERSATION,
];

/**
 * Run all tests
 * To be called from Google Apps Script environment
 */
function runNegotiationToggleTests() {
  const results = [];

  console.log('='.repeat(60));
  console.log('NEGOTIATION TOGGLE TESTS');
  console.log('='.repeat(60));

  for (const test of ALL_TESTS) {
    console.log(`\nRunning: ${test.name}`);
    console.log(`Description: ${test.description}`);
    console.log('-'.repeat(40));

    try {
      // Call the actual test function from code.gs
      const response = testAiEmailResponse(test.testData);

      if (response.error) {
        results.push({
          name: test.name,
          passed: false,
          error: response.error
        });
        console.log(`ERROR: ${response.error}`);
        continue;
      }

      // Validate the response
      const validation = test.validate(
        response.aiResponse || '',
        response.activeTypes || []
      );

      results.push({
        name: test.name,
        passed: validation.passed,
        errors: validation.errors,
        response: response.aiResponse?.substring(0, 200) + '...',
        activeTypes: response.activeTypes
      });

      if (validation.passed) {
        console.log('✓ PASSED');
      } else {
        console.log('✗ FAILED');
        validation.errors.forEach(e => console.log(`  - ${e}`));
      }

    } catch (error) {
      results.push({
        name: test.name,
        passed: false,
        error: error.message
      });
      console.log(`EXCEPTION: ${error.message}`);
    }
  }

  // Summary
  console.log('\n' + '='.repeat(60));
  console.log('TEST SUMMARY');
  console.log('='.repeat(60));

  const passed = results.filter(r => r.passed).length;
  const failed = results.filter(r => !r.passed).length;

  console.log(`Passed: ${passed}/${results.length}`);
  console.log(`Failed: ${failed}/${results.length}`);

  if (failed > 0) {
    console.log('\nFailed tests:');
    results.filter(r => !r.passed).forEach(r => {
      console.log(`  - ${r.name}`);
      if (r.error) console.log(`    Error: ${r.error}`);
      if (r.errors) r.errors.forEach(e => console.log(`    ${e}`));
    });
  }

  return results;
}

/**
 * Run a single test by name
 */
function runSingleTest(testName) {
  const test = ALL_TESTS.find(t => t.name === testName);
  if (!test) {
    console.log(`Test not found: ${testName}`);
    return null;
  }

  console.log(`Running: ${test.name}`);
  console.log(`Description: ${test.description}`);

  const response = testAiEmailResponse(test.testData);

  if (response.error) {
    return { passed: false, error: response.error };
  }

  const validation = test.validate(
    response.aiResponse || '',
    response.activeTypes || []
  );

  return {
    ...validation,
    response: response.aiResponse,
    activeTypes: response.activeTypes,
    negotiationState: response.negotiationState
  };
}

// Export for use in Google Apps Script
if (typeof module !== 'undefined') {
  module.exports = {
    ALL_TESTS,
    runNegotiationToggleTests,
    runSingleTest,
    analyzeResponse
  };
}
