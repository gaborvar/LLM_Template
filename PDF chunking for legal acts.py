
####################################################
# This function explores the PDF and separates it into chunks (sections) at logical boundaries.
# 
# It relies on common formatting standards of the EU's publications of draft and final Acts, for example the Official Journal.
# This heuristical approach limits the generalization of this code to any PDF or other source documents. 
# This procedural code could be replaced with document analysis e.g. by Azure AI Document Intelligence (aka Forms Recognizer )
# https://azure.microsoft.com/en-us/products/ai-services/ai-document-intelligence/
####################################################

from pdfminer.high_level import extract_pages    #  to examine PDF documents. Installation:  pip install pdfminer.six
from pdfminer.layout import LTContainer, LTChar, LTTextContainer


def extract_chunks(task_id, pdf_path, footer_pattern, paragraphpattern, top_margin, bottom_margin):

# First, collect statistics about the pattern of elements the PDF layout so we can recognise headings and their hierarchy. 
# With a histogram of the horizontal positions of all elements we could later infer if an element is centered which is a strong indication of a heading.

    bucket_count = 901      # granularity measure: number of buckets horizontally for the histogram of horizontal position of text elements. This can be finetuned. Set higher (narrower buckets) if too much body text is assumed to be heading.
    linelength_limit = 0.6 # if line is longer than this then it will not be considered a heading. Can be finetuned. Set lower if too much body text is assumed to be heading.
        # line length and granularity interplay: both are meant to exclude false positives when distinguishing heading from body text. 
        # try increasing one and decreasing the other for better identification of heading elements

    logging.info("Statistic collection in PDF for heading identification, parameters: number of buckets for position histogram: %s, linelength limit: %s",bucket_count, linelength_limit)

    center_positions = []

    for page_layout in extract_pages(pdf_path):
        page_width = page_layout.width  # Assuming page layout has width attribute
        
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                x0 = element.x0
                x1 = element.x1
                centerpos = (x0 + x1) / 2
                if (x1-x0) < page_layout.width/2:       # only keep short lines to avoid body text lines contaminating the heading statistics
                    center_positions.append(centerpos)  # issue: body text lines that happen to be centered at the same horizontal position would still be identified as heading

    # Normalize center positions to be a fraction of the page width
    center_positions_normalized = [pos / page_width for pos in center_positions]

    # Calculate histogram with buckets width of 1/n of the page width
    histogram, bin_edges = np.histogram(center_positions_normalized, bins=np.linspace(0, 1, bucket_count+1))  # Specifying 902, 901 buckets created. Better to pick an odd number so the middle will fall into a single bucket. 

    # Find the top X center positions (bin centers)
    top_X_bins = np.argsort(histogram)[-8:]         # 8 best bins retained
    top_X_centers = [(bin_edges[i] + bin_edges[i+1])/2 for i in top_X_bins]

    # Identify the two center positions close to a % of the page width
    centers_near_mid = [center for center in top_X_centers if abs(center - 0.55) < 0.15]  # Adjust threshold as needed. We assume that headings are centered between 40 to 70 % of page width

    # top_5_centers, centers_near_60
    # Assuming centers_near_60 are calculated 
    centers_near_mid_indices = [np.searchsorted(bin_edges, center) - 1 for center in centers_near_mid]  # create a list of indices also

    logging.info("PDF layout statistics: Top bins: %s,  Top centers: %s,  Indices near center of page: %s", top_X_bins, top_X_centers, centers_near_mid_indices)


    ##########################################
    # Separation of the PDF into chunks at headings starts here
    ##########################################

    chunks = []     # textual content of the PDF will be accumulated here. 
                    # Fields: text, heading, page number, embeddings
    current_heading = ""
    prev_consec_heading = ""
    pagenum = None
    text_buffer = ""    # Each chunk will be accumulated here as text. Heading and other text as well as paragraph marker also added
    font_names = []     # to collect all font names to see if font names are used other than italics or kurziv   

    search_results[task_id]['status']  = "Finished layout analysis of the document. Analysing formatting and text elements. Please wait."



    def alignedcenter(element, pagewd):         # Examine the text element and decide if it is a heading based on its horizontal position. Uses statistic created in the beginning.

        # todo: disregard if in footer
        nonlocal bin_edges, centers_near_mid_indices, linelength_limit  # in python: " each execution of extract_chunks creates a new instance of alignedcenter with its own environment, including the variables captured via nonlocal." so no risk of task interference through these variables
        elementpos_normalized = (element.x1+element.x0) / 2 / pagewd
        bin_index = np.digitize(elementpos_normalized, bin_edges) - 1       # Identify the horizonatal position bin that the position of the examined element falls into

        # next, check if the bin of the element is one of the bins that many headings belong to  
        alignedcenter = bin_index in centers_near_mid_indices and (element.x1-element.x0) < pagewd * linelength_limit   # Disregard line if longer than 51% of page. It is unlikely a heading.
        # todo: not protected from body text lines that happen to be centered around one of the heading centers, and whose length is small enough to assume it is a heading
        if alignedcenter: 
            logging.debug(f"Alignment testing of '{element.get_text().strip()}',  centerpos {elementpos_normalized},  bin_index {bin_index}, width: {element.x1-element.x0}")

        return alignedcenter



    # Recursive function to examine the formatting of the text element: boldness, italicness, size.
    # Assumes bold conservatively: assumes boldness only if all characters are bold and only if "Bold" (uppercase initial) is in the font name. 
        # Otherwise, does not declare it as bold. 
        # However, footnote reference (number) with small font is disregarded: cannot misrepresent boldness if only the reference is non-bold

    def check_and_update_font_properties(element):      
        nonlocal fontsize_min, fontsize_max, font_boldness, font_italicness, font_names, element_text   # also extracts text which  may be different from text of the get_text() of the higher level element

        if isinstance(element, LTChar):
            # Update global font size min/max
            fontsize_min = min(fontsize_min, element.size)
            fontsize_max = max(fontsize_max, element.size)

            if element.fontname not in font_names:
                font_names.append(element.fontname)

            # Update global font boldness:  'Bold' in font name - only if uppercase B
            if 'Bold' not in element.fontname and element.get_text().strip() and element.size >= 9:     #   Exclude small print to prevent footnote references from affecting boldness detection
                font_boldness = False
            
            if 'italic' not in element.fontname.lower() and element.get_text().strip():  # 'italic" in font name, case insensitive
                font_italicness = False
            # in Official Journal PDFs this will NOT catch italic text as the font name includes "Ital", not "Italic". 
            # However, this is not an issue as OJ PDFs do not include changes marked with bold italic (which should be excluded from heading identification). 
            # OJ journal PDFs can rely on center positioning of the heading, not only the font of the heading.  
            # (based on AI Act PDF from OJ)
            

            # Append text
            element_text += element.get_text()      # todo done: if no trailing space at the end of the element before the closing nl char we should add back. 

        elif hasattr(element, "get_text") and element.get_text() == "\n":   #   '\n' is not an LTChar so we have to handle them separately.
            element_text += " "             # Add a space, removing unnecessary line breaks. todo: if nl is the last char of the container then it may be better to add a nl to signify true paragraph end.
                                            # todo: Use a marker to signal that a nl was the last char of the element.
        elif isinstance(element, LTContainer):
            for child in element:
                check_and_update_font_properties(child)     # go one level deeper in the hierarchy
                #   todo: processing of an element is finished. We may be able to test here if text is yielded to element_text. If the last char is a marker for a replaced nl, then change the last char back to nl



    for page_layout in extract_pages(pdf_path):

        page_width = page_layout.width  # Assuming page layout has width attribute
        page_header_height =  page_layout.height * 0.92 if not top_margin else top_margin    # setting the top margin. Default is 8 % of the page
        page_footer_height =  page_layout.height * 0.08 if not bottom_margin else bottom_margin  # setting the bottom margin. Anything below this is considered a footer
        

        for element in page_layout:
            if element.y0 > page_header_height: 
                continue

            par_marker=""

            # Semi-Global variables
            fontsize_min = float('inf')     # to be updated in check_and_update_font_properties(element)
            fontsize_max = float('-inf')    # to be updated in check_and_update_font_properties(element)
            font_boldness = True        # to be updated in check_and_update_font_properties(element)
            font_italicness = True
            element_text = ""       # to be updated in check_and_update_font_properties(element)

            if hasattr(element, "get_text"):        # detect paragraph start with regex patterns
                element_level_text = element.get_text()     # this is different from extracting the text by going deep in the element and adding char by char

                if element.y1 < page_footer_height:     #  Check if the element is in the bottom margin area
                    # Remove footnote. Parse page number if a regex pattern is provided (currently null).
                    if footer_pattern:      # If a footer is specifid in the metadata. Should not span multiple elements. Currently null.
                        pgnum = re.search(footer_pattern, element_level_text) 
                        if pgnum:   # a page number search is provided a result 
                            if len(pgnum.groups()) >= 1 and pgnum.group(1):  # store page number if found by the first of the odd/even page footer patterns
                                pagenum = pgnum.group(1)
                            elif len(pgnum.groups()) > 1:   # store the page number if the other of the two footer patterns found it 
                                pagenum = pgnum.group(2)
                    continue    # do not proceed with the current element. Skip to the next one (within the same footer or on the next page )

                if re.match(paragraphpattern, element_level_text.strip()):  #  eliminated a repeated call to get_text
                    par_marker = element_level_text
                    logging.info("Para marker: '%s', len: %s", element_level_text.strip(), len(text_buffer))

            check_and_update_font_properties(element)  # updates fontsize_min, fontsize_max, font_boldness, element_text

            # check if it contains text other than small print.
            if element_text and fontsize_max >= 8.4:    
                # In the AI Act PDFs paragraph marker (number) may be 8.49 points so threashold must be smaller than 8.49. 
                # However, in the Data Act PDF footnotes are identified (and excluded) on the basis of small font size. Threshold must be larger than 8.5294
                # Unresolved conflict. Perhaps a PDF specific setting in the metadata may address this issue but it is not a general solution.
                # potential resolution: switch to Document Intelligence, an Azure service to extract PDF content https://documentintelligence.ai.azure.com/studio
                
                if not element_text.endswith((' ', '\n')):  # standardize trailing white space: add space if none
                    # logging.info(f"Adding trailing space to '{element_text}'")                
                    element_text = element_text + ' '

                # Check if the element is a paragraph marker

                if par_marker:
                    # If current buffer size is long enough, split here: store and empty the buffer 
                    if len(text_buffer) > 1400: # shortened from 2000 to keep chunks more readable
                        chunks.append({"text": text_buffer, "heading": current_heading, "page": pagenum, "embeddings":""})  # embeddings is a placeholder for later
                        text_buffer = ""
                    # Add the paragraph marker whether or not the buffer is flushed
                    text_buffer += element_text

                else:

                    if (        # check whether it looks like a heading
                        element_text.strip()
                        and (fontsize_min > 13  
                             or (font_boldness == True and font_italicness == False)    # assume heading if bold but not italic
                             or (font_italicness == True and font_boldness == False)    # assume heading if not bold but italic. todo: only if not in footer
                                # in pre-final versions of acts changes are often signalled with bold italic. 
                                # Thus, bold italic text must not be considered a heading - unless overriding heuristics suggest, see below
                             or alignedcenter(element, page_width)  # assume heading if positioned in the horizontal center of the page. (no notion of text alignment in PDF or PDFMiner) 
                             ) 
                        # compare text regardless of white space. Assume header only if the two text extractions match.
                        and re.sub(r'\s+', ' ', element_level_text).strip()  == re.sub(r'\s+', ' ', element_text).strip()  
                        and len(element_text) > 6   # short text is not a heading
                        and len(element_text) < 90  # long text is not a heading
                        ):

                        # logging.info("Heading: '%s', midpos:  %s", element_text, (element.x0 + element.x1)/2)
                        if len(text_buffer) > 500:      # If the chunk is more than 500 chars then flush and start new
                            chunks.append({"text": text_buffer, "heading": current_heading, "page": pagenum, "embeddings":""})
                            text_buffer = ""


                        current_heading = prev_consec_heading + element_text    # We allow two headings after each other to be treated as one. 
                        prev_consec_heading = element_text      # todo: consider adding " -- " between the two consec headings. Currently only a space is added, see above

                    else:       # does not look like a heading
                        
                        prev_consec_heading = ""

                        if len(text_buffer) > 5000:      # Prevent overly long chunks even if no heading or paragraph is detected. If the chunk is more than X chars then flush and start new
                            chunks.append({"text": text_buffer, "heading": current_heading, "page": pagenum, "embeddings":""})
                            text_buffer = ""


                    text_buffer += element_text

        # End of page. 

        search_results[task_id]['status']  = "Chunking, currently under heading:  '" + current_heading + f"'  Pagenumber if detected: {pagenum}"  # todo: check if current_heading is null
        # logging.info("A page is finished. Status stored for the client progress message.")


    # Add the last buffer if it exists
    chunks.append({"text": text_buffer, "heading": current_heading, "page": pagenum, "embeddings":""})
    text_buffer = ""
    
    return chunks, font_names

