from pptx import Presentation                                                                                                                                                              
from pptx.dml.color import ColorFormat, RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt
import sys

Grouped_Triggers = ['UMH', 'FWS', 'Scripture Reading', 'Prayer for Illumination', "The Lord’s Prayer"] 
Standalone_Triggers = ['Passing of the Peace', 'Rev.', 'Prelude', 'Postlude', 'The Children’s Moment', 'Offering Our Gifts', 'Offertory', 'Sending Forth', 'Proclamation of God’s Word'] 

def extract_data( prs ):
    '''extract_data takes a given Presentation object and returns a list of lists containing the slide number, slide "type", and slide content of that object'''
    data = []
    Flag = "Standalone"
    slide_no = 1

    for slide in prs.slides:
        
        raw_data = []
        raw_data.append( slide_no )
        cleaned_text = ""

        for shape in slide.shapes:
            raw_text = ""

            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    raw_text = ''.join( run.text.strip() ) + ' '
                    if any( substr in raw_text for substr in Grouped_Triggers) == True:
                        Flag = "Grouped"
                    if any( substr in raw_text for substr in Standalone_Triggers) == True:
                        Flag = "Standalone"
                    if raw_text != '':
                        cleaned_text += raw_text
        raw_data.append( Flag )
        raw_data.append( cleaned_text )
        data.append( raw_data )
        slide_no += 1

    return data

def Main():
    if len( sys.argv ) != 2:
        if len( sys.argv ) < 2:
            print( "Error: No path to .pptx file provided" )
            exit( 1 )
        elif len( sys.argv ) > 2:
            print( "Error: To many arguments provided [Expected 1 argument, " + str( len( sys.argv ) - 1 ) + " arguments provided]" )
            exit( 1 )
    else:
        file_path = sys.argv[1]
        if file_path[-5::1] != ".pptx":
            print( "Error: expected .pptx file" )
            exit( 1 )
        else:
            presentation = Presentation( file_path )
            presentation.slide_width = Inches(16)  # Width in inches
            presentation.slide_height = Inches(9)  # Height in inches
            blank_slide_layout = presentation.slide_layouts[6]
            
            data = extract_data( presentation )
            for raw_data in data:
                print( raw_data )

if __name__ == "__main__":
    Main()
