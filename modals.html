<div id="tutorialOverlay"></div>
<div id="tutorialTooltip">
  <div class="tut-header-stripe">
    <div style="display:flex; align-items:center; gap:10px;">
      <span class="material-icons-round" style="color:rgba(255,255,255,0.85); font-size:18px;">school</span>
      <span style="color:white; font-size:11px; font-weight:700; letter-spacing:1px; text-transform:uppercase;">Interactive Tour</span>
    </div>
    <div style="display:flex; align-items:center; gap:10px;">
      <div class="tut-step-dots" id="tutStepDots"></div>
      <div class="tut-step-badge" id="tutProgressText">STEP 1 OF 10</div>
    </div>
  </div>
  <div class="tut-progress-bar-track">
    <div class="tut-progress-bar-fill" id="tutProgressBar" style="width:10%;"></div>
  </div>
  <div class="tut-body">
    <div class="tut-title"><span id="tutTitleText"></span></div>
    <div class="tut-desc" id="tutDescText"></div>
    <div id="tutActionHint"></div>
  </div>
  <div class="tut-footer">
    <button class="tut-btn-skip" onclick="endTutorial()">
      <span class="material-icons-round" style="font-size:14px;">close</span> Exit Tour
    </button>
    <div class="tut-btn-row">
      <button class="tut-btn-prev" id="tutPrevBtn" onclick="prevTutorialStep()">
        <span class="material-icons-round" style="font-size:14px;">arrow_back</span> Back
      </button>
      <button class="tut-btn-next" id="tutNextBtn" onclick="nextTutorialStep()">
        Next <span class="material-icons-round" style="font-size:14px;">arrow_forward</span>
      </button>
    </div>
  </div>
</div>

<div class="toast-notification" id="statusToast">
  <div class="toast-dot"></div>
  <span id="toastMessage">Updating matrix parameters...</span>
</div>

<div class="universal-backdrop-mask" id="calendarModal" onclick="toggleCalendarModal(false)">
  <div class="modal-surface-container" onclick="event.stopPropagation()">
    <div class="modal-header-row">
      <div class="modal-title-text" id="modalTargetDateLabel">Filed Leaves Matrix</div>
      <button class="modal-close-btn" onclick="toggleCalendarModal(false)">×</button>
    </div>
    <div class="modal-roster-list" id="modalRosterContentContainer"></div>
  </div>
</div>

<div class="universal-backdrop-mask" id="profileModal" onclick="toggleProfileModal(false)">
  <div class="modal-surface-container" onclick="event.stopPropagation()">
    <div class="modal-header-row">
      <div class="modal-title-text" id="modalProfileNameLabel">Employee Core Registry Profile</div>
      <button class="modal-close-btn" onclick="toggleProfileModal(false)">×</button>
    </div>
    <div class="modal-profile-scroll-zone">
      <div class="demographics-profile-grid" id="modalProfileGridBody"></div>
    </div>
  </div>
</div>

<div class="universal-backdrop-mask" id="sproutComplianceModal">
  <div class="modal-surface-container sprout-surface">
    <div class="sprout-icon-shield">⚠️</div>
    <div class="sprout-headline">Filing Logged Successfully</div>
    <div class="sprout-description">
      This transaction has been synced to the database tracker. Please remember to log this request in your official portal to align payroll:
      <br>
      <a href="https://sso.sprout.ph/realms/saperium/protocol/openid-connect/auth?client_id=SproutSSO&redirect_uri=https%3A%2F%2Fsaperium.hrhub.ph%2F&response_type=code%20id_token&scope=openid%20profile&state=OpenIdConnect.AuthenticationProperties%3DVJFMv5UmYIbOCrct_Jrnny2Zm3CPpETFRFxK8gAupBBGQb7noyeVIEzX4xpSYuLhPVKdAda_8uYcvZcrBdTZV2IkV2NVBtyfGmxXXrSzOiQUQiy5uZlQf_kK2alUIMCp0YNSV4HK5C4295BImXnkc_YyHowV0Vh1i3i-2I-UubahMRjOukdt7J_LpYfp76FX8EyeVA&response_mode=form_post&nonce=639195507541121783.MTEyN2IyZjQtMjBkNC00ODA5LTliZjEtNGNmM2Q3MDdkYjkzNjJiNjcyYzAtY2Q2OC00MmY1LWI2NGItYTRhZDNlMjZiNWU0&x-client-SKU=ID_NET461&x-client-ver=5.3.0.0" target="_blank" class="sprout-redirect-link">Sprout Portal →</a>
    </div>
    <button class="sprout-confirm-btn" onclick="toggleSproutModal(false)">Acknowledge & Close</button>
  </div>
</div>

<div class="universal-backdrop-mask" id="deleteConfirmModal" onclick="toggleDeleteConfirmModal(false)">
  <div class="modal-surface-container sprout-surface" onclick="event.stopPropagation()">
    <div class="warning-icon-shield">
      <span class="material-icons-round" style="font-size:30px;">delete_sweep</span>
    </div>
    <div class="sprout-headline">Confirm Leave Deletion</div>
    <div class="sprout-description" id="deleteModalWarningTextBody">
      Warning: You are about to clear the active leave transaction details for this resource.
    </div>
    <div class="warning-modal-button-controls-row">
      <button class="warning-cancel-btn" onclick="toggleDeleteConfirmModal(false)">Cancel</button>
      <button class="warning-confirm-danger-btn" id="executeDeleteBtnNode" onclick="executeLeaveDeletionTransaction()">Delete Leave</button>
    </div>
  </div>
</div>
<div class="universal-backdrop-mask" id="reassignmentModal">
  <div class="modal-surface-container sprout-surface">
    <div class="warning-icon-shield" style="background:#FEF3C7; border-color:#F59E0B; color:#D97706;">
      <span class="material-icons-round" style="font-size:30px;">group_remove</span>
    </div>
    <div class="sprout-headline">Coverage Gap Detected</div>
    <div class="sprout-description" id="reassignModalDescText" style="text-align:left;">
      A critical coverage gap has been detected. Both the Main POC and Backup are on leave for the requested date. Please select an available team member to temporarily cover this task.
    </div>
    <div id="reassignTaskContainer" style="margin-top:15px; text-align:left;"></div>
    <div class="warning-modal-button-controls-row" style="margin-top: 15px;">
      <button class="sprout-confirm-btn" onclick="applyReassignments()">Resolve & Continue</button>
    </div>
  </div>
</div>
