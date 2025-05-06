import streamlit as st
import json
import urllib.parse
from streamlit_drawable_canvas import st_canvas
import time
import numpy as np

def titleAndLogo(title):
    st.markdown(f"<h1 style='text-align: center;'>{title}<h1>", unsafe_allow_html=True)
    
    # Only load URL parameters once per session
    if 'url_params_loaded' not in st.session_state:
        try:
            load_from_query_params()
            # Mark that we've loaded the URL parameters
            st.session_state['url_params_loaded'] = True
        except Exception as e:
            st.error(f"Error loading form data: {e}")
    return

def load_from_query_params():
    """Load form data from individual URL parameters"""
    try:
        query_params = st.query_params
        
        # If we have query parameters (excluding reserved ones)
        valid_params = [k for k in query_params.keys() if k not in ['streamlit_app_url', 'data']]
        
        if valid_params:
            st.info("Loading form data from shared link...")
            
            # Transfer all parameters to session state
            loaded_fields = 0
            for key in valid_params:
                value = query_params[key]
                
                # Check if this is a numeric parameter
                is_numeric_param = (
                    # Keys that should be floats
                    key == 'date' or 
                    key.startswith('total_design_flow_ms') or
                    key.startswith('total_supply_design_flow_ms') or
                    key.startswith('design_flow_') or
                    key.startswith('anem_')
                )
                
                # Check if this is an integer parameter
                is_int_param = (
                    key == 'num_canopies' or
                    key.startswith('quantity') or
                    key.startswith('ksaQuantity')
                )
                
                # Special handling for date fields
                if key == 'date':
                    try:
                        import datetime
                        date_obj = datetime.date.fromisoformat(value)
                        st.session_state[key] = date_obj
                    except:
                        st.session_state[key] = value
                # Special handling for boolean values
                elif value.lower() in ['true', 'false']:
                    st.session_state[key] = value.lower() == 'true'
                # Special handling for integer values
                elif is_int_param:
                    try:
                        st.session_state[key] = int(float(value))
                    except:
                        # If conversion fails, try to parse as boolean
                        if value.lower() in ['true', 'false']:
                            st.session_state[key] = value.lower() == 'true'
                        else:
                            # Otherwise keep as string
                            st.session_state[key] = value
                # Special handling for float values
                elif is_numeric_param:
                    try:
                        st.session_state[key] = float(value)
                    except:
                        # Keep as string if conversion fails
                        st.session_state[key] = value
                # Text inputs
                elif key in ['client', 'projectName', 'projectNumber', 'engineers', 'comments']:
                    st.session_state[key] = str(value)
                # Everything else
                else:
                    st.session_state[key] = value
                
                loaded_fields += 1
            
            if loaded_fields > 0:
                st.success(f"Successfully loaded {loaded_fields} fields from shared link")
    except Exception as e:
        st.error(f"Error loading query parameters: {str(e)}")
        import traceback
        st.code(traceback.format_exc())

def getGenInfo():
    # Ensure text input session state values are strings
    for key in ['client', 'projectName', 'projectNumber', 'engineers']:
        if key in st.session_state and not isinstance(st.session_state[key], str):
            st.session_state[key] = str(st.session_state[key])
    
    col1, col2 = st.columns(2)
    with col1:
        # st.markdown("<h2 style='text-align:center; margin-bottom:-40px;'>Client</h2>", unsafe_allow_html=True)
        client = st.text_input('Client', key='client')
    with col2:
        # st.markdown("<h2 style='text-align:center;margin-bottom:-40px;'>Project Name</h2>", unsafe_allow_html=True)
        projectName = st.text_input('Project Name', key='projectName')
    with col1:
        # st.markdown("<h2 style='text-align:center;margin-bottom:-40px;'>Project Number</h2>", unsafe_allow_html=True)
        projectNumber= st.text_input('Project Number', key='projectNumber')
    with col2:
        # st.markdown("<h2 style='text-align:center;margin-bottom:-40px;'>Date Of Visit</h2>", unsafe_allow_html=True)
        DateOfVisit = st.date_input('Date', key='date')
    # st.markdown("<h2 style='text-align:center;margin-bottom:-40px;'>Engineer(s)</h2>", unsafe_allow_html=True)
    engineers = st.text_input('Engineers', key='engineers')
    
    return {
        'client': client,
        'project_name': projectName,
        'project_number': projectNumber,
        'date_of_visit': DateOfVisit,
        'engineers': engineers
    }

def get_comments():
    # Ensure comments session state value is a string
    if 'comments' in st.session_state and not isinstance(st.session_state['comments'], str):
        st.session_state['comments'] = str(st.session_state['comments'])
        
    comments = st.text_area('Comments (Enter multiple comments separated by "/")', key='comments')
    # Split by slash instead of newline to allow multiple comments in one line
    comments_split = comments.split('/')
    # Trim whitespace from each comment
    comments_split = [comment.strip() for comment in comments_split if comment.strip()]
    return comments_split

def get_sign():
    """Get signature input and ensure it persists between reruns"""
    st.write("Signed By Engineer(s)")
    
    # Initialize signature data in session state if not already present
    if 'signature_data' not in st.session_state:
        st.session_state.signature_data = None
    
    # Create the canvas widget
    canvas_result = st_canvas(
        stroke_width=5,
        stroke_color="#000000",
        background_color="#FFFFFF",
        height=150,
        width=700,
        drawing_mode="freedraw",
        key="signature_canvas"
    )
    
    # If a new drawing is provided, save it to session state
    if canvas_result.image_data is not None:
        # Compare with session state to see if it's different (new drawing)
        st.session_state.signature_data = canvas_result.image_data
    
    # If nothing was drawn but we have saved data, restore it
    elif st.session_state.signature_data is not None:
        # This doesn't directly update the canvas itself, but ensures 
        # the image_data is returned for processing
        canvas_result.image_data = st.session_state.signature_data
        
        # Show the stored signature as an image below the canvas
        # This ensures the user can always see their signature
        st.image(st.session_state.signature_data, caption="Your saved signature", width=700)
    
    return canvas_result

def create_shareable_link():
    """Create a shareable link with form data as individual URL parameters"""
    # Collect essential form data
    form_data = {}
    for key, value in st.session_state.items():
        # Skip canvas data, functions, and empty values
        try:
            # Skip numpy arrays, complex objects, and empty values
            if ('canvas' not in key and 
                not callable(value) and 
                value is not None and 
                not isinstance(value, np.ndarray) and
                key not in ['base_url', 'signature_data']):
                
                # Skip empty strings
                if isinstance(value, str) and value == "":
                    continue
                    
                # For date objects, convert to string
                if hasattr(value, 'isoformat'):
                    form_data[key] = value.isoformat()
                else:
                    # Make sure it's serializable
                    json.dumps({key: value})
                    form_data[key] = value
        except:
            # Skip values that can't be serialized or compared
            continue
    
    # Get the base URL from session state
    base_url = st.session_state.get('base_url', 'http://localhost:8501')
    base_url = base_url.rstrip('/')
    
    # Build URL with individual parameters
    params = []
    for key, value in form_data.items():
        # URL encode each value
        encoded_value = urllib.parse.quote(str(value))
        params.append(f"{key}={encoded_value}")
    
    # Join all parameters with &
    query_string = "&".join(params)
    
    # Create complete URL
    full_url = f"{base_url}?{query_string}"
    
    return full_url

def display_shareable_link():
    """Display a shareable link with the form data"""
    # Add a URL input field to let users specify their base URL
    base_url = st.text_input(
        "Your app's URL (where this form is hosted):",
        value="https://haltonservice-uk.streamlit.app/",
        key="base_url",
        help="Enter the URL where your app is hosted. For local development, use http://localhost:8501"
    )
    
    if st.button("Generate Shareable Link"):
        with st.spinner("Creating shareable link..."):
            full_url = create_shareable_link()
            
            st.success("✅ Shareable link created!")
            
            # Display the full URL in a text area for easy copying
            st.text_area(
                "Copy this complete URL and share with others:",
                value=full_url,
                height=100
            )
            
            st.info("""
            **Instructions:**
            1. Copy the complete link above
            2. Share it with anyone who needs to access your form data
            3. When they open the link, all fields will be pre-filled with your data
            """)
            
            # If URL is very long, provide a download alternative
            if len(full_url) > 2000:
                st.warning("⚠️ This URL is quite long and might not work in all browsers. Consider using the download option below instead.")
                
                # Add download option
                form_data = {}
                for key, value in st.session_state.items():
                    if 'canvas' not in key and not callable(value):
                        # Handle date objects
                        if hasattr(value, 'isoformat'):
                            form_data[key] = value.isoformat()
                        else:
                            try:
                                json.dumps({key: value})
                                form_data[key] = value
                            except:
                                form_data[key] = str(value)
                
                json_data = json.dumps(form_data)
                
                st.download_button(
                    label="💾 Download Form Data",
                    data=json_data,
                    file_name=f"form_data_{st.session_state.get('projectNumber', 'project')}.json",
                    mime="application/json"
                )
    