#pragma once

// Resource compiler compatibility.
#ifndef IDC_STATIC
#define IDC_STATIC (-1)
#endif

// -------------------------
// Dialog IDs
// -------------------------
#define IDD_MAIN                 100
#define IDD_SOFTVOICE_SAY        IDD_MAIN   // alias (optional)

// -------------------------
// Main controls (top-level)
// -------------------------
#define IDC_ENGINE_PATH          1001
#define IDC_ENGINE_BROWSE        1002

#define IDC_TEXT                1003
#define IDC_TEXT_INPUT          IDC_TEXT    // alias (if your .rc uses IDC_TEXT_INPUT)

#define IDC_OPEN_TEXT            1004
#define IDC_OPEN_FILE            IDC_OPEN_TEXT  // alias (if your .rc uses IDC_OPEN_FILE)

#define IDC_SPEAK                1005
#define IDC_STOP                 1006
#define IDC_SAVE_WAV             1007
#define IDC_STATUS               1008

// -------------------------
// Combo boxes
// -------------------------
#define IDC_VOICE                1010
#define IDC_VARIANT              1011
#define IDC_SMODE                1012   // speaking mode
#define IDC_SPEAKING_MODE        IDC_SMODE  // alias

#define IDC_GENDER               1013
#define IDC_INTSTYLE             1014
#define IDC_VMODE                1015
#define IDC_GLOT                 1016

// -------------------------
// Numeric edits + spins
// -------------------------
#define IDC_RATE                 1100
#define IDC_RATE_SPIN            1101

#define IDC_PITCH                1110
#define IDC_PITCH_SPIN           1111

#define IDC_INFLECTION           1120
#define IDC_INFLECTION_SPIN      1121

#define IDC_PAUSE                1130
#define IDC_PAUSE_SPIN           1131

#define IDC_PERTURB              1140
#define IDC_PERTURB_SPIN         1141

#define IDC_VFACTOR              1150
#define IDC_VFACTOR_SPIN         1151

// Voicing / frication / aspiration gains (naming matches softvoice-say.cpp)
#define IDC_AVBIAS               1160
#define IDC_AVBIAS_SPIN          1161

#define IDC_AFBIAS               1170
#define IDC_AFBIAS_SPIN          1171

#define IDC_AHBIAS               1180
#define IDC_AHBIAS_SPIN          1181
