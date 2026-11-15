"""Book-specific adapters used by the FBD GUI.

Each adapter keeps one book's parsing and PPT generation rules out of the GUI
layer.  The GUI should collect files and report progress; adapters decide how a
book's FBD files are converted into the MD runs expected by its PPT generator.
"""

