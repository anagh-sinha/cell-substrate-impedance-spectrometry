// extractor.c
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <windows.h>
#include <commdlg.h>

#define MAX_FILES      100
#define MAX_FREQS      300
#define MAX_LOOPS      100
#define MAX_LABEL_LEN  128
#define MAX_LINE_LEN   512

// --- globals ----------------------------------------------------------------
char *filenames[MAX_FILES];
char  file_labels[MAX_FILES][MAX_LABEL_LEN];
int   total_files = 0;

double freqs[MAX_FREQS];
int    freq_count = 0;

// raw_primary[file][loop][freqIdx], raw_secondary[...]:
double raw_primary[MAX_FILES][MAX_LOOPS][MAX_FREQS];
double raw_secondary[MAX_FILES][MAX_LOOPS][MAX_FREQS];
// how many loops seen so far per file/freq
int    loop_count[MAX_FILES][MAX_FREQS];
// maximum loops for each file (across all freqs)
int    max_loops[MAX_FILES];

// --- find or add frequency, returns its index -------------------------------
int find_or_add_freq(double f) {
    for (int i = 0; i < freq_count; i++) {
        if (fabs(freqs[i] - f) < 1e-3) return i;
    }
    freqs[freq_count] = f;
    return freq_count++;
}

// --- for sorting frequencies -----------------------------------------------
int cmp_order(const void *a, const void *b) {
    int ia = *(int*)a;
    int ib = *(int*)b;
    if (freqs[ia] < freqs[ib]) return -1;
    if (freqs[ia] > freqs[ib]) return  1;
    return 0;
}

// --- parse one file’s List_Meas_Result section -----------------------------
void process_file(const char *filename, int fileIdx) {
    FILE *f = fopen(filename, "r");
    if (!f) return;

    char line[MAX_LINE_LEN];
    int in_section = 0;

    while (fgets(line, sizeof(line), f)) {
        if (!in_section) {
            if (strstr(line, "List_Meas_Result")) {
                in_section = 1;
                fgets(line, sizeof(line), f);  // skip header
            }
            continue;
        }

        double freq, primary, second;
        char unit1[16], unit2[16];
        // this sscanf matches the original format you had;
        // it pulls out: freq, primary, primary‐unit, second, second‐unit
        int fields = sscanf(
            line,
            "%*d %*d %*d %*s %*s %*s %lf %*f %*s %lf %s",
            &freq, &primary, unit1, &second, unit2
        );
        if (fields == 5) {
            // convert “Mohm” → “kohm”
            if (strcmp(unit1, "Mohm") == 0) primary *= 1000.0;
            if (strcmp(unit2, "Mohm") == 0) second  *= 1000.0;

            int idx  = find_or_add_freq(freq);
            int loop = loop_count[fileIdx][idx]++;
            raw_primary  [fileIdx][loop][idx] = primary;
            raw_secondary[fileIdx][loop][idx] = second;

            if (loop+1 > max_loops[fileIdx])
                max_loops[fileIdx] = loop+1;
        }
    }

    fclose(f);
}

// --- print the full tab-delimited table to any FILE* (stdout or file) -------
void output_table(FILE *out) {
    // 1) build sorted index array
    int order[MAX_FREQS];
    for (int i = 0; i < freq_count; i++) order[i] = i;
    qsort(order, freq_count, sizeof(int), cmp_order);

    // 2) header
    fprintf(out, "Freq(kHz)");
    for (int f = 0; f < total_files; f++) {
        for (int L = 0; L < max_loops[f]; L++) {
            fprintf(
              out,
              "\tPrimary_%s_Loop%d(kohm)"
              "\tSecondary_%s_Loop%d(kohm)",
              file_labels[f], L+1,
              file_labels[f], L+1
            );
        }
    }
    fprintf(out, "\n");

    // 3) one line per frequency
    for (int r = 0; r < freq_count; r++) {
        int idx = order[r];
        fprintf(out, "%.3f", freqs[idx]);

        for (int f = 0; f < total_files; f++) {
            for (int L = 0; L < max_loops[f]; L++) {
                if (L < loop_count[f][idx]) {
                    fprintf(out, "\t%.3f", raw_primary  [f][L][idx]);
                    fprintf(out, "\t%.3f", raw_secondary[f][L][idx]);
                } else {
                    // no data for that loop → blank cell
                    fprintf(out, "\t");
                    fprintf(out, "\t");
                }
            }
        }
        fprintf(out, "\n");
    }
}

// --- user picks files → process → show table → save dialog → write file ---
int main() {
    OPENFILENAME ofn;
    char filebuf[8192] = "";
    char outpath[MAX_PATH] = "";

    // 1) select one or more TXT inputs
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize  = sizeof(ofn);
    ofn.lpstrFilter  = "Text Files\0*.txt\0All Files\0*.*\0";
    ofn.lpstrFile    = filebuf;
    ofn.nMaxFile     = sizeof(filebuf);
    ofn.Flags        = OFN_ALLOWMULTISELECT | OFN_EXPLORER | OFN_FILEMUSTEXIST;
    ofn.lpstrTitle   = "Select Input TXT Files";

    if (!GetOpenFileName(&ofn)) {
        printf("Input selection cancelled.\n");
        return 1;
    }

    // expand the buffer into filenames[]
    char *p = filebuf;
    char dir[MAX_PATH];
    strcpy(dir, p);
    p += strlen(p) + 1;
    if (*p == '\0') {
        filenames[ total_files++ ] = strdup(dir);
    } else {
        while (*p) {
            char full[MAX_PATH];
            snprintf(full, sizeof(full), "%s\\%s", dir, p);
            filenames[ total_files++ ] = strdup(full);
            p += strlen(p) + 1;
        }
    }

    // derive simple labels (basename without “.txt”) for each file
    for (int i = 0; i < total_files; i++) {
        char *b = strrchr(filenames[i], '\\');
        if (!b) b = strrchr(filenames[i], '/');
        b = b ? b + 1 : filenames[i];
        char *dot = strrchr(b, '.');
        int   len = dot ? (int)(dot - b) : (int)strlen(b);
        if (len >= MAX_LABEL_LEN) len = MAX_LABEL_LEN-1;
        strncpy(file_labels[i], b, len);
        file_labels[i][len] = '\0';
    }

    // 2) process each file
    for (int i = 0; i < total_files; i++) {
        process_file(filenames[i], i);
    }

    // 3) show table on console
    output_table(stdout);

    // 4) ask where to save
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.lpstrFilter = "Text Files\0*.txt\0";
    ofn.lpstrFile   = outpath;
    ofn.nMaxFile    = sizeof(outpath);
    ofn.Flags       = OFN_OVERWRITEPROMPT;
    ofn.lpstrTitle  = "Save Output TXT File";

    if (!GetSaveFileName(&ofn)) {
        printf("Save cancelled.\n");
        return 1;
    }

    // 5) write the exact same table into that file
    FILE *fo = fopen(outpath, "w");
    if (fo) {
        output_table(fo);
        fclose(fo);
        MessageBox(NULL, "TXT file successfully written.", "Done", MB_OK);
    } else {
        MessageBox(NULL, "Failed to open output file.", "Error", MB_ICONERROR);
    }

    return 0;
}
