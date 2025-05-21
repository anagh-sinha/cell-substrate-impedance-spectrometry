#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <windows.h>
#include <commdlg.h>

#define MAX_FILES 100
#define MAX_LINE_LEN 512
#define MAX_FREQS 300
#define MAX_LOOPS 100
#define MAX_LABEL_LEN 10000

// Data structure for each unique frequency and each file's loops
typedef struct {
    double freq;
    double primary[MAX_FILES][MAX_LOOPS];
    double secondary[MAX_FILES][MAX_LOOPS];
    int count[MAX_FILES];    // loops per file
} FrequencyData;

static FrequencyData freq_data[MAX_FREQS];
static int freq_count = 0;
static int total_files = 0;
static char *file_labels[MAX_FILES];

// Compare frequencies for sorting
int compare_freq(const void* a, const void* b) {
    double fa = ((FrequencyData*)a)->freq;
    double fb = ((FrequencyData*)b)->freq;
    return (fa > fb) - (fa < fb);
}

// Find or add frequency index
int find_or_add_freq(double freq) {
    for (int i = 0; i < freq_count; i++) {
        if (fabs(freq_data[i].freq - freq) < 1e-6) return i;
    }
    // initialize new entry
    freq_data[freq_count].freq = freq;
    for (int f = 0; f < total_files; f++)
        freq_data[freq_count].count[f] = 0;
    return freq_count++;
}

// Process one input file into freq_data[file_idx]
void process_file(const char* filename, int file_idx) {
    FILE* f = fopen(filename, "r");
    if (!f) {
        fprintf(stderr, "Failed to open %s\n", filename);
        return;
    }
    char line[MAX_LINE_LEN];
    int in_results = 0;
    while (fgets(line, sizeof(line), f)) {
        if (strstr(line, "List_Meas_Result")) {
            in_results = 1;
            fgets(line, sizeof(line), f); // skip header
            continue;
        }
        if (!in_results) continue;
        double freq, prim, sec;
        char unit1[16] = {0}, unit2[16] = {0};
        int fields = sscanf(line,
            "%*d %*d %*d %*s %*s %*s %lf %*f %*s %lf %s %lf %s",
            &freq, &prim, unit1, &sec, unit2
        );
        if (fields == 5) {
            // Normalize into kOhm
            if (strcmp(unit1, "Mohm") == 0) {
                prim *= 1000.0;             // MΩ → kΩ
            } else if (strcmp(unit1, "Ohm") == 0) {
                prim /= 1000.0;             // Ω → kΩ
            }
            if (strcmp(unit2, "Mohm") == 0) {
                sec *= 1000.0;
            } else if (strcmp(unit2, "Ohm") == 0) {
                sec /= 1000.0;
            }
            // Accept kohm, Mohm, or Ohm inputs
            if ((strcmp(unit1, "kohm") == 0 || strcmp(unit1, "Mohm") == 0 || strcmp(unit1, "Ohm") == 0) &&
                (strcmp(unit2, "kohm") == 0 || strcmp(unit2, "Mohm") == 0 || strcmp(unit2, "Ohm") == 0)) {
                int idx = find_or_add_freq(freq);
                int loop_idx = freq_data[idx].count[file_idx]++;
                if (loop_idx < MAX_LOOPS) {
                    freq_data[idx].primary[file_idx][loop_idx]   = prim;
                    freq_data[idx].secondary[file_idx][loop_idx] = sec;
                }
            }
        }
    }
    fclose(f);
}

// Determine max loops per file
typedef int MaxLoopsArray[MAX_FILES];
void get_max_loops(MaxLoopsArray max_loops) {
    for (int f = 0; f < total_files; f++)
        max_loops[f] = 0;
    for (int i = 0; i < freq_count; i++) {
        for (int f = 0; f < total_files; f++) {
            if (freq_data[i].count[f] > max_loops[f])
                max_loops[f] = freq_data[i].count[f];
        }
    }
}

// Print TAB-separated table to given FILE*
void output_table(FILE* out) {
    MaxLoopsArray max_loops;
    get_max_loops(max_loops);
    // Header
    fprintf(out, "Freq(kHz)");
    for (int f = 0; f < total_files; f++) {
        for (int l = 0; l < max_loops[f]; l++) {
            fprintf(out, "\tPrimary_%s_Loop%d(kohm)\tSecondary_%s_Loop%d(kohm)",
                    file_labels[f], l+1,
                    file_labels[f], l+1);
        }
    }
    fprintf(out, "\n");
    // Sort
    qsort(freq_data, freq_count, sizeof(FrequencyData), compare_freq);
    // Rows
    for (int i = 0; i < freq_count; i++) {
        fprintf(out, "%.3f", freq_data[i].freq);
        for (int f = 0; f < total_files; f++) {
            for (int l = 0; l < max_loops[f]; l++) {
                if (l < freq_data[i].count[f])
                    fprintf(out, "\t%.3f\t%.3f",
                            freq_data[i].primary[f][l],
                            freq_data[i].secondary[f][l]);
                else
                    fprintf(out, "\t\t");
            }
        }
        fprintf(out, "\n");
    }
}

// Write TXT file
void write_txt(const char* out_filename) {
    FILE* out = fopen(out_filename, "w");
    if (!out) { fprintf(stderr, "Failed to create %s\n", out_filename); return; }
    output_table(out);
    fclose(out);
}

int main() {
    OPENFILENAME ofn;
    char file_buffer[8192] = "";
    char output_file[MAX_PATH] = "";
    char* filenames[MAX_FILES];

    // Select input files
    ZeroMemory(&ofn, sizeof(ofn)); ofn.lStructSize = sizeof(ofn);
    ofn.lpstrFilter = "Text Files\0*.txt\0All Files\0*.*\0";
    ofn.lpstrFile   = file_buffer; ofn.nMaxFile = sizeof(file_buffer);
    ofn.Flags       = OFN_ALLOWMULTISELECT | OFN_EXPLORER | OFN_FILEMUSTEXIST;
    ofn.lpstrTitle  = "Select Input TXT Files";
    if (!GetOpenFileName(&ofn)) return printf("Input cancelled.\n"), 1;

    // Parse filenames & labels
    char dir[MAX_PATH]; char* p = file_buffer;
    strcpy(dir, p); p += strlen(p)+1;
    total_files = (*p=='\0')?1:0;
    if (total_files==1) filenames[0]=dir;
    else while(*p && total_files<MAX_FILES) {
        char full[MAX_PATH]; sprintf(full, "%s\\%s", dir, p);
        filenames[total_files++] = strdup(full); p+=strlen(p)+1;
    }
    for(int i=0;i<total_files;i++){
        char* b = strrchr(filenames[i], '\\');
        b = b ? b+1 : filenames[i];
        char tmp[MAX_LABEL_LEN];
        strncpy(tmp, b, MAX_LABEL_LEN-1);
        tmp[MAX_LABEL_LEN-1] = '\0';
        char* d = strrchr(tmp, '.'); if (d) *d = '\0';
        file_labels[i] = strdup(tmp);
    }

    // Process files
    for (int i = 0; i < total_files; i++)
        process_file(filenames[i], i);

    // Print to console
    output_table(stdout);

    // Save file
    ZeroMemory(&ofn, sizeof(ofn)); ofn.lStructSize = sizeof(ofn);
    ofn.lpstrFilter = "Text Files\0*.txt\0All Files\0*.*\0";
    ofn.lpstrFile   = output_file; ofn.nMaxFile = sizeof(output_file);
    ofn.Flags       = OFN_OVERWRITEPROMPT; ofn.lpstrTitle = "Save Output TXT File";
    if (!GetSaveFileName(&ofn)) return printf("Save cancelled.\n"), 1;

    write_txt(output_file);
    MessageBox(NULL, "Data successfully exported to TXT.", "Done", MB_OK);
    return 0;
}