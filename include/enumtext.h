#ifndef ENUMTEXT_H
#define ENUMTEXT_H

#include <QString>

namespace Docx {

enum class WD_PARAGRAPH_ALIGNMENT
{
    LEFT = 0,
    CENTER,
    RIGHT,
    JUSTIFY,
    DISTRIBUTE,
    JUSTIFY_MED,
    JUSTIFY_HI,
    JUSTIFY_LOW,
    THAI_JUSTIFY
};

enum class WD_TABLE_ALIGNMENT
{
    LEFT = 0,
    CENTER,
    RIGHT,
    BOTH,
    DISTRIBUTE,
    MEDIUMKASHIDA,
    HIGHKASHIDA,
    LOWKASHIDA,
    THAIDISTRIBUTE
};

enum class WD_UNDERLINE
{
    None = -1,
    SINGLE,
    WORDS,
    DOUBLE,
    DOTTED,
    THICK,
    DASH,
    DOT_DASH,
    DOT_DOT_DASH,
    WAVY,
    DOTTED_HEAVY,
    DASH_HEAVY,
    DOT_DASH_HEAVY,
    DOT_DOT_DASH_HEAVY,
    WAVY_HEAVY,
    DASH_LONG,
    WAVY_DOUBLE,
    DASH_LONG_HEAVY,

};

enum class WD_BREAK
{
    LINE,
    PAGE,
    COLUMN,
    LINE_CLEAR_LEFT,
    LINE_CLEAR_RIGHT,
    LINE_CLEAR_ALL
};

enum class WD_FONTSIZE
{
    SIZE0 = 84,
    SIZE00 = 72,
    SIZE1 = 52,
    SIZE01 = 48,
    SIZE2 = 44,
    SIZE02 = 36,
    SIZE3 = 32,
    SIZE03 = 30,
    SIZE4 = 28,
    SIZE04 = 24,
    SIZE5 = 21,
    SIZE05 = 18,
    SIZE6 = 15,
    SIZE06 = 13,
    SIZE7 = 11,
    SIZE8 = 10,
};

enum DW_IMAGELAYOUT
{
    InLineWithText,
    Square,
    Tight,
    Through,
    TopAnBottom,
    BehindText,
    InFrontOfText
};

QString underLineToString(WD_UNDERLINE underline);

QString paragraphAlignToString(WD_PARAGRAPH_ALIGNMENT align);

}// namespace Docx

#endif// ENUMTEXT_H
