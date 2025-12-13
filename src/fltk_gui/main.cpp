#include <FL/Fl.H>
#include "FltkMainWindow.h"

int main(int argc, char **argv) {
    // Set visual style
    Fl::scheme("gtk+");
    
    // Title: Excel Data Processing Tool
    FltkMainWindow window(1200, 800, "Excel\xE6\x95\xB0\xE6\x8D\xAE\xE5\xA4\x84\xE7\x90\x86\xE5\xB7\xA5\xE5\x85\xB7-\xE6\xA8\xAA\xE6\xA2\x81\xE4\xB8\x8B\xE6\x96\x99\xE4\xB8\x93\xE4\xB8\x9A\xE7\xBB\x84");
    window.show(argc, argv);
    return Fl::run();
}
