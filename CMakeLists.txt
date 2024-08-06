cmake_minimum_required(VERSION 3.12)
project(DOCX)

set(BUILD_WITH_QT6 FALSE CACHE BOOL "Enable (experimental) Qt6 support")
if(BUILD_WITH_QT6)
    set(QT_VERSION_MAJOR 6)
    set(QT_MIN_VERSION 6.4.0)
    set(QT_VERSION_BASE "Qt6")
    set(QT_VERSION_BASE_LOWER "qt6")
    set(CMAKE_CXX_STANDARD 17)
    set(CMAKE_CXX_STANDARD_REQUIRED ON)
else()
    set(QT_VERSION_MAJOR 5)
    set(QT_MIN_VERSION 5.15.2)
    set(QT_VERSION_BASE "Qt5")
    set(QT_VERSION_BASE_LOWER "qt5")
    set(HAS_KDE_QT5_PDF_TRANSFORM_FIX FALSE CACHE BOOL "Using KDE's Qt 5.15 fork with the PDF brush transform fix")
    set(HAS_KDE_QT5_SMALL_CAPS_FIX FALSE CACHE BOOL "Using KDE's Qt 5.15 fork with the QFont stretch fix")
endif()

find_package(${QT_VERSION_BASE} COMPONENTS Core Gui Widgets Xml REQUIRED)

set(CMAKE_AUTOMOC ON)

set(CMAKE_DEBUG_POSTFIX "d")
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_CURRENT_LIST_DIR}/bin)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_CURRENT_LIST_DIR}/bin)
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_CURRENT_LIST_DIR}/lib)

include(GenerateExportHeader)

set(DOCX_INCLUDE_DIR ${CMAKE_CURRENT_LIST_DIR}/include)
set(DOCX_INCLUDE
    ${DOCX_INCLUDE_DIR}/document.h
    ${DOCX_INCLUDE_DIR}/length.h
    ${DOCX_INCLUDE_DIR}/shape.h
    ${DOCX_INCLUDE_DIR}/table.h
    ${DOCX_INCLUDE_DIR}/text.h
    ${DOCX_INCLUDE_DIR}/enumtext.h
)

set(DOCX_SOURCES 
    src/image/image.cpp 
    src/opc/parts/corepropertiespart.cpp 
    src/opc/part.cpp 
    src/opc/packuri.cpp 
    src/opc/opcpackage.cpp 
    src/opc/coreproperties.cpp 
    src/opc/oxml.cpp 
    src/opc/physpkgreader.cpp 
    src/opc/packagereader.cpp 
    src/opc/contenttypemap.cpp 
    src/opc/serializedrelationships.cpp 
    src/opc/rel.cpp 
    src/opc/packagewriter.cpp 
    src/opc/physpkgwriter.cpp 
    src/oxml/xmlchemy.cpp 
    src/oxml/oxmltable.cpp 
    src/oxml/oxmlshape.cpp 
    src/oxml/oxmltext.cpp 
    src/parts/documentpart.cpp 
    src/parts/imagepart.cpp 
    src/shape.cpp 
    src/enumtext.cpp 
    src/length.cpp
    src/document.cpp 
    src/shared.cpp 
    src/text.cpp 
    src/package.cpp    
    src/table.cpp 
)

set(DOCX_HEADERS 
    src/mage/image.h 
    src/pc/parts/corepropertiespart.h 
    src/pc/part.h 
    src/pc/packuri.h 
    src/pc/opcpackage.h 
    src/pc/coreproperties.h 
    src/pc/oxml.h 
    src/pc/constants.h  
    src/pc/physpkgreader.h 
    src/pc/packagereader.h 
    src/pc/contenttypemap.h 
    src/pc/serializedrelationships.h 
    src/pc/rel.h 
    src/pc/packagewriter.h 
    src/pc/physpkgwriter.h    
    src/xml/oxmltext.h 
    src/xml/xmlchemy.h 
    src/xml/oxmltable.h 
    src/xml/oxmlshape.h 
    src/arts/imagepart.h 
    src/arts/documentpart.h 
    src/hared.h 
    src/ackage.h 
)

add_library(docxlib SHARED ${DOCX_INCLUDE} ${DOCX_SOURCES} ${DOCX_HEADERS})

GENERATE_EXPORT_HEADER(
  docxlib
  BASE_NAME DOCX
  EXPORT_FILE_NAME docx_global.h
)

set(DOCX_INCLUDE ${DOCX_INCLUDE} ${CMAKE_CURRENT_BINARY_DIR}/docx_global.h)

target_include_directories(docxlib SYSTEM PUBLIC
  ${${QT_VERSION_BASE}Gui_PRIVATE_INCLUDE_DIRS}
  ${DOCX_INCLUDE_DIR}
  ${CMAKE_CURRENT_BINARY_DIR}
)

target_link_libraries(docxlib 
  ${QT_VERSION_BASE}::Core 
  ${QT_VERSION_BASE}::Widgets 
  ${QT_VERSION_BASE}::Gui 
  ${QT_VERSION_BASE}::Xml
  ${QT_VERSION_BASE}::GuiPrivate
)

if(BUILD_TEST)
    include(CTest)
    add_subdirectory(tests)
endif()

install(TARGETS docxlib
  LIBRARY DESTINATION lib  
  ARCHIVE DESTINATION lib  
  RUNTIME DESTINATION bin  
  PUBLIC_HEADER DESTINATION include  
)

install(FILES ${DOCX_INCLUDE} DESTINATION include)