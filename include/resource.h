// resource.h for softvoice-say

#pragma once

#ifndef IDC_STATIC
#define IDC_STATIC (-1)
#endif

// Aliases to keep .rc and .cpp IDs compatible
#ifndef IDD_SOFTVOICE_SAY
#define IDD_SOFTVOICE_SAY IDD_MAIN
#endif
#ifndef IDC_TEXT_INPUT
#define IDC_TEXT_INPUT IDC_TEXT
#endif
#ifndef IDC_OPEN_FILE
#define IDC_OPEN_FILE IDC_OPEN_TEXT
#endif
#ifndef IDC_SPEAKING_MODE
#define IDC_SPEAKING_MODE IDC_SMODE
#endif

#define IDD_MAIN 101

#define IDC_ENGINE_PATH    1001
#define IDC_ENGINE_BROWSE  1002

#define IDC_TEXT           1003
#define IDC_OPEN_TEXT      1004
#define IDC_SPEAK          1005
#define IDC_STOP           1006
#define IDC_SAVE_WAV       1007

#define IDC_VOICE          1010
#define IDC_VARIANT        1011
#define IDC_SMODE          1012

#define IDC_RATE           1013
#define IDC_RATE_SPIN      1014
#define IDC_PITCH          1015
#define IDC_PITCH_SPIN     1016
#define IDC_INFLECTION     1017
#define IDC_INFLECTION_SPIN 1018
#define IDC_PAUSE          1019
#define IDC_PAUSE_SPIN     1020

#define IDC_PERTURB        1021
#define IDC_PERTURB_SPIN   1022
#define IDC_VFACTOR        1023
#define IDC_VFACTOR_SPIN   1024
#define IDC_AVBIAS         1025
#define IDC_AVBIAS_SPIN    1026
#define IDC_AFBIAS         1027
#define IDC_AFBIAS_SPIN    1028
#define IDC_AHBIAS         1029
#define IDC_AHBIAS_SPIN    1030

#define IDC_INTSTYLE       1031
#define IDC_VMODE          1032
#define IDC_GENDER         1033
#define IDC_GLOT           1034

#define IDC_STATUS         1050
