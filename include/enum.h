#ifndef ENUMTEXT_H
#define ENUMTEXT_H

#include <QString>

namespace Docx {

/*
Specifies the color specification scheme.

http://msdn.microsoft.com/en-us/library/office/ff864912(v=office.15).aspx
*/
enum MsoColorType
{
    RGB = 1, // Color is specified by an |RGBColor| value.
    THEME = 2, // Color is one of the preset theme colors.
    AUTO = 101, // Color is determined automatically by the application.
};

/*
Indicates the Office theme color, one of those shown in the color gallery on the
formatting ribbon.

http://msdn.microsoft.com/en-us/library/office/ff860782(v=office.15).aspx
*/
enum MsoThemeColorIndex
{
    NOT_THEME_COLOR = 0, // Indicates the color is not a theme color.
    ACCENT_1 = 5, // Specifies the Accent 1 theme color.
    ACCENT_2 = 6, // Specifies the Accent 2 theme color.
    ACCENT_3 = 7, // Specifies the Accent 3 theme color.
    ACCENT_4 = 8, // Specifies the Accent 4 theme color.
    ACCENT_5 = 9, // Specifies the Accent 5 theme color."
    ACCENT_6 = 10, // Specifies the Accent 6 theme color.
    BACKGROUND_1 = 14, // Specifies the Background 1 theme color.
    BACKGROUND_2 = 16, // Specifies the Background 2 theme color.
    DARK_1 = 1, // Specifies the Dark 1 theme color.
    DARK_2 = 3, // Specifies the Dark 2 theme color.
    FOLLOWED_HYPERLINK = 12, // Specifies the theme color for a clicked hyperlink.
    HYPERLINK = 11, // Specifies the theme color for a hyperlink.
    LIGHT_1 = 2, // Specifies the Light 1 theme color.
    LIGHT_2 = 4, // Specifies the Light 2 theme color.
    TEXT_1 = 13, // Specifies the Text 1 theme color.
    TEXT_2 = 15, // Specifies the Text 2 theme color.
};

/*
Specifies one of the three possible header/footer definitions for a section.

https://docs.microsoft.com/en-us/office/vba/api/word.wdheaderfooterindex
*/
enum WdHeaderFooterIndex
{
    PRIMARY, // Header for odd pages or all if no even header.
    FIRST_PAGE, // Header for first page of section.
    EVEN_PAGE // Header for even pages of recto/verso section.
};


/*
Specifies the page layout orientation.

http://msdn.microsoft.com/en-us/library/office/ff837902.aspx
*/
enum WdOrientation
{
    PORTRAIT, // Portrait orientation.
    LANDSCAPE // Landscape orientation.
};

/*
Specifies the start type of a section break.

http://msdn.microsoft.com/en-us/library/office/ff840975.aspx
*/
enum WdSectionStart
{
    CONTINUOUS, // Continuous section break.
    NEW_COLUMN, // New column section break.
    NEW_PAGE, // New page section break.
    EVEN_PAGE, // Even pages section break.
    ODD_PAGE, // Section begins on next odd page.
};

/*
Enumerations related to DrawingML shapes in WordprocessingML files.
Corresponds to WdInlineShapeType enumeration.

http://msdn.microsoft.com/en-us/library/office/ff192587.aspx.
*/
enum WdInlineShapeType
{
    CHART = 12
    LINKED_PICTURE = 4
    PICTURE = 3
    SMART_ART = 15
    NOT_IMPLEMENTED = -6
};


/*
Specifies a built-in Microsoft Word style.

http://msdn.microsoft.com/en-us/library/office/ff835210.aspx
*/
enum WdBuiltinStyle
{
    BLOCK_QUOTATION = -85, // Block Text.
    BODY_TEXT = -67, // Body Text.
    BODY_TEXT_2 = -81, // Body Text 2.
    BODY_TEXT_3 = -82, // Body Text 3.
    BODY_TEXT_FIRST_INDENT = -78, // Body Text First Indent.
    BODY_TEXT_FIRST_INDENT_2 = -79, // Body Text First Indent 2.
    BODY_TEXT_INDENT = -68, // Body Text Indent.
    BODY_TEXT_INDENT_2 = -83, // Body Text Indent 2.
    BODY_TEXT_INDENT_3 = -84, // Body Text Indent 3.
    BOOK_TITLE = -265, // Book Title.
    CAPTION = -35, // Caption.
    CLOSING = -64, // Closing.
    COMMENT_REFERENCE = -40, // Comment Reference.
    COMMENT_TEXT = -31, // Comment Text.
    DATE = -77, // Date.
    DEFAULT_PARAGRAPH_FONT = -66, // Default Paragraph Font.
    EMPHASIS = -89, // Emphasis.
    ENDNOTE_REFERENCE = -43, // Endnote Reference.
    ENDNOTE_TEXT = -44, // Endnote Text.
    ENVELOPE_ADDRESS = -37, // Envelope Address.
    ENVELOPE_RETURN = -38, // Envelope Return.
    FOOTER = -33, // Footer.
    FOOTNOTE_REFERENCE = -39, // Footnote Reference.
    FOOTNOTE_TEXT = -30, // Footnote Text.
    HEADER = -32, // Header.
    HEADING_1 = -2, // Heading 1.
    HEADING_2 = -3, // Heading 2.
    HEADING_3 = -4, // Heading 3.
    HEADING_4 = -5, // Heading 4.
    HEADING_5 = -6, // Heading 5.
    HEADING_6 = -7, // Heading 6.
    HEADING_7 = -8, // Heading 7.
    HEADING_8 = -9, // Heading 8.
    HEADING_9 = -10, // Heading 9.
    HTML_ACRONYM = -96, // HTML Acronym.
    HTML_ADDRESS = -97, // HTML Address.
    HTML_CITE = -98, // HTML Cite.
    HTML_CODE = -99, // HTML Code.
    HTML_DFN = -100, // HTML Definition.
    HTML_KBD = -101, // HTML Keyboard.
    HTML_NORMAL = -95, // Normal (Web).
    HTML_PRE = -102, // HTML Preformatted.
    HTML_SAMP = -103, // HTML Sample.
    HTML_TT = -104, // HTML Typewriter.
    HTML_VAR = -105, // HTML Variable.
    HYPERLINK = -86, // Hyperlink.
    HYPERLINK_FOLLOWED = -87, // Followed Hyperlink.
    INDEX_1 = -11, // Index 1.
    INDEX_2 = -12, // Index 2.
    INDEX_3 = -13, // Index 3.
    INDEX_4 = -14, // Index 4.
    INDEX_5 = -15, // Index 5.
    INDEX_6 = -16, // Index 6.
    INDEX_7 = -17, // Index 7.
    INDEX_8 = -18, // Index 8.
    INDEX_9 = -19, // Index 9.
    INDEX_HEADING = -34, // Index Heading
    INTENSE_EMPHASIS = -262, // Intense Emphasis.
    INTENSE_QUOTE = -182, // Intense Quote.
    INTENSE_REFERENCE = -264, // Intense Reference.
    LINE_NUMBER = -41, // Line Number.
    LIST = -48, // List.
    LIST_2 = -51, // List 2.
    LIST_3 = -52, // List 3.
    LIST_4 = -53, // List 4.
    LIST_5 = -54, // List 5.
    LIST_BULLET = -49, // List Bullet.
    LIST_BULLET_2 = -55, // List Bullet 2.
    LIST_BULLET_3 = -56, // List Bullet 3.
    LIST_BULLET_4 = -57, // List Bullet 4.
    LIST_BULLET_5 = -58, // List Bullet 5.
    LIST_CONTINUE = -69, // List Continue.
    LIST_CONTINUE_2 = -70, // List Continue 2.
    LIST_CONTINUE_3 = -71, // List Continue 3.
    LIST_CONTINUE_4 = -72, // List Continue 4.
    LIST_CONTINUE_5 = -73, // List Continue 5.
    LIST_NUMBER = -50, // List Number.
    LIST_NUMBER_2 = -59, // List Number 2.
    LIST_NUMBER_3 = -60, // List Number 3.
    LIST_NUMBER_4 = -61, // List Number 4.
    LIST_NUMBER_5 = -62, // List Number 5.
    LIST_PARAGRAPH = -180, // List Paragraph.
    MACRO_TEXT = -46, // Macro Text.
    MESSAGE_HEADER = -74, // Message Header.
    NAV_PANE = -90, // Document Map.
    NORMAL = -1, // Normal.
    NORMAL_INDENT = -29, // Normal Indent.
    NORMAL_OBJECT = -158, // Normal (applied to an object).
    NORMAL_TABLE = -106, // Normal (applied within a table).
    NOTE_HEADING = -80, // Note Heading.
    PAGE_NUMBER = -42, // Page Number.
    PLAIN_TEXT = -91, // Plain Text.
    QUOTE = -181, // Quote.
    SALUTATION = -76, // Salutation.
    SIGNATURE = -65, // Signature.
    STRONG = -88, // Strong.
    SUBTITLE = -75, // Subtitle.
    SUBTLE_EMPHASIS = -261, // Subtle Emphasis.
    SUBTLE_REFERENCE = -263, // Subtle Reference.
    TABLE_COLORFUL_GRID = -172, // Colorful Grid.
    TABLE_COLORFUL_LIST = -171, // Colorful List.
    TABLE_COLORFUL_SHADING = -170, // Colorful Shading.
    TABLE_DARK_LIST = -169, // Dark List.
    TABLE_LIGHT_GRID = -161, // Light Grid.
    TABLE_LIGHT_GRID_ACCENT_1 = -175, // Light Grid Accent 1.
    TABLE_LIGHT_LIST = -160, // Light List.
    TABLE_LIGHT_LIST_ACCENT_1 = -174, // Light List Accent 1.
    TABLE_LIGHT_SHADING = -159, // Light Shading.
    TABLE_LIGHT_SHADING_ACCENT_1 = -173, // Light Shading Accent 1.
    TABLE_MEDIUM_GRID_1 = -166, // Medium Grid 1.
    TABLE_MEDIUM_GRID_2 = -167, // Medium Grid 2.
    TABLE_MEDIUM_GRID_3 = -168, // Medium Grid 3.
    TABLE_MEDIUM_LIST_1 = -164, // Medium List 1.
    TABLE_MEDIUM_LIST_1_ACCENT_1 = -178, // Medium List 1 Accent 1.
    TABLE_MEDIUM_LIST_2 = -165, // Medium List 2.
    TABLE_MEDIUM_SHADING_1 = -162, // Medium Shading 1.
    TABLE_MEDIUM_SHADING_1_ACCENT_1 = -176, // Medium Shading 1 Accent 1.
    TABLE_MEDIUM_SHADING_2 = -163, // Medium Shading 2.
    TABLE_MEDIUM_SHADING_2_ACCENT_1 = -177, // Medium Shading 2 Accent 1.
    TABLE_OF_AUTHORITIES = -45, // Table of Authorities.
    TABLE_OF_FIGURES = -36, // Table of Figures.
    TITLE = -63, // Title.
    TOAHEADING = -47, // TOA Heading.
    TOC_1 = -20, // TOC 1.
    TOC_2 = -21, // TOC 2.
    TOC_3 = -22, // TOC 3.
    TOC_4 = -23, // TOC 4.
    TOC_5 = -24, // TOC 5.
    TOC_6 = -25, // TOC 6.
    TOC_7 = -26, // TOC 7.
    TOC_8 = -27, // TOC 8.
    TOC_9 = -28, // TOC 9.
};


/*
Specifies one of the four style types: paragraph, character, list, or table.

http://msdn.microsoft.com/en-us/library/office/ff196870.aspx
*/
enum WdStyleType
{
    CHARACTER = 2, // Character style
    LIST = 4, // List style.
    PARAGRAPH = 1, //Paragraph style.
    TABLE = 3 // Table style
};


/*
Specifies the vertical alignment of text in one or more cells of a table.

https://msdn.microsoft.com/en-us/library/office/ff193345.aspx
*/
enum WdCellVerticalAlignment
{
    TOP = 0, // Text is aligned to the top border of the cell.
    CENTER = 1, // Text is aligned to the center of the cell.
    BOTTOM = 3, // Text is aligned to the bottom border of the cell.
    BOTH = 101, //This is an option in the OpenXml spec, but not in Word itself. It's not clear what Word behavior this setting produces. If you find out please let us know and we'll update this documentation. Otherwise, probably best to avoid this option.
}


/*
Specifies the rule for determining the height of a table row

https://msdn.microsoft.com/en-us/library/office/ff193620.aspx
*/
enum WdRowHeightRule
{
    AUTO, // The row height is adjusted to accommodate the tallest value in the row.
    AT_LEAST, // The row height is at least a minimum specified value.
    EXACTLY // The row height is an exact value.
};

/*
Specifies table justification type.

http://office.microsoft.com/en-us/word-help/HV080607259.aspx
*/
enum WdRowAlignment
{
    LEFT, // Left-aligned.
    CENTER, // Center-aligned.
    RIGHT, // Right-aligned.
};

/*
Specifies the direction in which an application orders cells in the specified table or row.

http://msdn.microsoft.com/en-us/library/ff835141.aspx
*/
enum WdTableDirection
{
    LTR, // The table or row is arranged with the first column in the leftmost position.
    RTL, // The table or row is arranged with the first column in the rightmost position.
};

/*
Specifies paragraph justification type.
*/
enum WdAlignPararaph
{
    LEFT = 0, // Left-aligned.
    CENTER = 1, // Center-aligned.
    RIGHT = 2, // Right-aligned.
    JUSTIFY = 3, // Fully justified.
    DISTRIBUTE = 4, // Paragraph characters are distributed to fill entire width of paragraph.
    JUSTIFY_MED =  5, // Justified with a medium character compression ratio.
    JUSTIFY_HI = 7, // Justified with a high character compression ratio.
    JUSTIFY_LOW = 8, // Justified with a low character compression ratio.
    THAI_JUSTIFY = 9,// Justified according to Thai formatting layout.
};


/*
Corresponds to WdBreakType enumeration.

http://msdn.microsoft.com/en-us/library/office/ff195905.aspx.
*/
enum WdBreakType
{
    COLUMN = 8
    LINE = 6
    LINE_CLEAR_LEFT = 9
    LINE_CLEAR_RIGHT = 10
    LINE_CLEAR_ALL = 11  // added for consistency, not in MS version
    PAGE = 7
    SECTION_CONTINUOUS = 3
    SECTION_EVEN_PAGE = 4
    SECTION_NEXT_PAGE = 2
    SECTION_ODD_PAGE = 5
    TEXT_WRAPPING = 11
};


/*
Specifies a standard preset color to apply.
Used for font highlighting and perhaps other applications.

https://msdn.microsoft.com/EN-US/library/office/ff195343.aspx
*/
enum WdColorIndex
{
    INHERITED = -1, // Color is inherited from the style hierarchy.
    AUTO = 0, // Automatic color. Default; usually black.
    BLACK = 1, // Black color.
    BLUE = 2, // Blue color.
    BRIGHT_GREEN = 4, // Bright green color.
    DARK_BLUE = 9, // Dark blue color.
    DARK_RED = 13, // Dark red color.
    DARK_YELLOW = 14, // Dark yellow color.
    GRAY_25 = 16, // 25% shade of gray color.
    GRAY_50 = 15, // 50% shade of gray color.
    GREEN = 11, // Green color.
    PINK = 5, // Pink color.
    RED = 6, // Red color.
    TEAL = 10, // Teal color.
    TURQUOISE = 3, // Turquoise color.
    VIOLET = 12, // Violet color.
    WHITE = 8, // White color.
    YELLOW = 7, // Yellow color.
};


/*
Specifies a line spacing format to be applied to a paragraph.

http://msdn.microsoft.com/en-us/library/office/ff844910.aspx
*/
enum WdLineSpacing
{
    SINGLE = 0, // Single spaced (default).
    ONE_POINT_FIVE = 1, // Space-and-a-half line spacing.
    DOUBLE = 2, // Double spaced.
    AT_LEAST = 3,// Minimum line spacing is specified amount. Amount is specified separately.
    EXACTLY = 4,// Line spacing is exactly specified amount. Amount is specified separately.
    MULTIPLE = 5,// Line spacing is specified as multiple of line heights. Changing font size will change line spacing proportionately.
};


/*
Specifies the tab stop alignment to apply.

https://msdn.microsoft.com/EN-US/library/office/ff195609.aspx
*/
enum WdTabAlignment
{
    LEFT = 0, // Left-aligned.
    CENTER = 1, // Center-aligned.
    RIGHT = 2, // Right-aligned.
    DECIMAL = 3, // Decimal-aligned.
    BAR = 4, // Bar-aligned.
    LIST = 6, // List-aligned. (deprecated)
    CLEAR = 101, // Clear an inherited tab stop.
    END = 102, // Right-aligned.  (deprecated)
    NUM = 103, // Left-aligned.  (deprecated)
    START = 104, // Left-aligned.  (deprecated)
};

/*
Specifies the character to use as the leader with formatted tabs.

https://msdn.microsoft.com/en-us/library/office/ff845050.aspx
*/
enum WdTabLeader
{
    SPACES = 0, // Spaces. Default.
    DOTS = 1, // Dots.
    DASHES = 2, // Dashes.
    LINES = 3, // Double lines.
    HEAVY = 4, // A heavy line.
    MIDDLE_DOT = 5, // A vertically-centered dot.
};


/*
Specifies the style of underline applied to a run of characters.

http://msdn.microsoft.com/en-us/library/office/ff822388.aspx
*/
enum WdUnderline
{
    INHERITED = -1, // Inherit underline setting from containing paragraph.")
    NONE = 0,// No underline.\n\nThis setting overrides any inherited underline value, so can be used to remove underline from a run that inherits underlining from its containing paragraph. Note this is not the same as assigning |None| to Run.underline. |None| is a valid assignment value, but causes the run to inherit its underline value. Assigning `WD_UNDERLINE.NONE` causes underlining to be unconditionally turned off.",
    SINGLE = 1,// A single line.\n\nNote that this setting is write-only in the sense that |True| (rather than `WD_UNDERLINE.SINGLE`) is returned for a run having this setting.
    WORDS = 2, // Underline individual words only.
    DOUBLE = 3, // A double line.
    DOTTED = 4, // Dots.
    THICK = 6, // A single thick line.
    DASH = 7, // Dashes.
    DOT_DASH = 9, // Alternating dots and dashes.
    DOT_DOT_DASH = 10, // An alternating dot-dot-dash pattern.
    WAVY = 11, // A single wavy line.
    DOTTED_HEAVY = 20, // Heavy dots.
    DASH_HEAVY = 23, // Heavy dashes.
    DOT_DASH_HEAVY = 25, // Alternating heavy dots and heavy dashes.
    DOT_DOT_DASH_HEAVY =  26, // An alternating heavy dot-dot-dash pattern.
    WAVY_HEAVY = 27, // A heavy wavy line.
    DASH_LONG = 39, // Long dashes.
    WAVY_DOUBLE = 43, // A double wavy line.
    DASH_LONG_HEAVY = 55, // Long heavy dashes.
};


}// namespace Docx

#endif// ENUMTEXT_H
