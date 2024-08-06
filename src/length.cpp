#include "length.h"
#include <math.h>

namespace Docx {

const int EMUS_PER_INCH = 914400;
const int EMUS_PER_CM = 360000;
const int EMUS_PER_MM = 36000;
const int EMUS_PER_PX = 12700;
const int EMUS_PER_TWIP = 635;

const int UNITS_PER_POINT = 100;

Length::Length() {}

Length::Length(int emu) : m_value(emu) { m_isEmpty = false; }

float Length::cm() const { return (float) m_value / (float) EMUS_PER_CM; }

int Length::emu() const { return m_value; }

float Length::inches() const { return (float) m_value / (float) EMUS_PER_INCH; }

float Length::mm() const { return (float) m_value / (float) EMUS_PER_MM; }

int Length::px() const { return (int) (round((float) m_value / (float) EMUS_PER_PX) + 0.1); }

int Length::twips() const { return (int) (round((float) m_value / (float) EMUS_PER_TWIP)); }

bool Length::isEmpty() const { return m_isEmpty; }

Length Inches::emus(float inches)
{
    int emu = int(inches * EMUS_PER_INCH);
    return Length(emu);
}

Length Cm::emus(float cm)
{
    int emu = int(cm * EMUS_PER_CM);
    return Length(emu);
}

Length Emu::emus(float emu) { return Length(emu); }

Length Mm::emus(float mm)
{
    int emu = int(mm * EMUS_PER_MM);
    return Length(emu);
}

Length Pt::emus(float pts)
{
    int units = int(pts * UNITS_PER_POINT);
    return Length(units);
}

Length Px::emus(float px)
{
    int emu = int(px * EMUS_PER_PX);
    return Length(emu);
}

Length Twips::emus(float twips)
{
    int emu = int(twips * EMUS_PER_TWIP);
    return Length(emu);
}

}// namespace Docx
