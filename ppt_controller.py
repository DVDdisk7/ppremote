import win32com.client
import os
import tempfile
import base64
import hashlib
import pythoncom
import glob

class PPTController:
    def __init__(self):
        self.temp_dir = tempfile.gettempdir()
        self.last_state = None
        self.last_pres_name = ""
        self.last_slide_idx = -1
    def _get_app(self):
        # We must call CoInitialize in the thread using COM
        pythoncom.CoInitialize()
        try:
            return win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception:
            return None

    def start_show(self):
        app = self._get_app()
        if app and app.Presentations.Count > 0:
            app.ActivePresentation.SlideShowSettings.Run()

    def stop_show(self):
        app = self._get_app()
        if app:
            try:
                app.SlideShowWindows(1).View.Exit()
            except Exception:
                pass

    def next_slide(self):
        app = self._get_app()
        if app:
            try:
                app.SlideShowWindows(1).View.Next()
            except Exception:
                pass

    def prev_slide(self):
        app = self._get_app()
        if app:
            try:
                app.SlideShowWindows(1).View.Previous()
            except Exception:
                pass

    def toggle_blank(self):
        app = self._get_app()
        if app:
            try:
                view = app.SlideShowWindows(1).View
                # State 3 is black screen, 1 is running, 4 is white screen, 5 is done
                if view.State == 3:
                    view.State = 1
                else:
                    view.State = 3
            except Exception:
                pass

    def goto_slide(self, index):
        app = self._get_app()
        if app:
            try:
                app.SlideShowWindows(1).View.GotoSlide(index)
            except Exception:
                pass

    def get_state(self):
        app = self._get_app()
        state = {
            "is_running": False,
            "current_slide": 0,
            "total_slides": 0,
            "notes": "",
            "current_image": None,
            "next_image": None
        }
        
        if not app:
            return state
            
        try:
            if app.SlideShowWindows.Count > 0:
                state["is_running"] = True
                ss = app.SlideShowWindows(1)
                view = ss.View
                pres = ss.Presentation
                
                current_idx = view.CurrentShowPosition
                pres_name = pres.Name
                is_blank = (view.State == 3 or view.State == 4)
                
                # PowerPoint can sometimes return 0 for CurrentShowPosition during rapid transitions.
                # If this happens, we return the last known valid state to avoid UI glitches.
                if current_idx <= 0 and self.last_state and self.last_pres_name == pres_name:
                    return self.last_state
                
                state["current_slide"] = current_idx
                state["total_slides"] = pres.Slides.Count
                state["is_blank"] = is_blank
                state["pres_name"] = pres_name
                
                if self.last_state and self.last_pres_name == pres_name and self.last_slide_idx == current_idx:
                    state["notes"] = self.last_state.get("notes", "")
                    state["current_image"] = self.last_state.get("current_image")
                    state["next_image"] = self.last_state.get("next_image")
                    return state
                
                # Get notes
                if 1 <= current_idx <= pres.Slides.Count:
                    try:
                        slide = pres.Slides(current_idx)
                        if slide.HasNotesPage:
                            notes_slide = slide.NotesPage
                            notes_text = []
                            for shape in notes_slide.Shapes:
                                if shape.HasTextFrame:
                                    if shape.TextFrame.HasText:
                                        txt = shape.TextFrame.TextRange.Text.strip()
                                        if txt and not txt.isdigit():
                                            notes_text.append(txt)
                            # The actual notes are usually the longest text
                            if notes_text:
                                state["notes"] = max(notes_text, key=len)
                            else:
                                state["notes"] = ""
                    except Exception as e:
                        print("Error getting notes:", e)
                
                # Calculate fast export dimensions preserving aspect ratio
                if not hasattr(self, 'slide_width'):
                    self.slide_width = pres.PageSetup.SlideWidth
                    self.slide_height = pres.PageSetup.SlideHeight
                
                scale_w = 800
                scale_h = int(800 * (self.slide_height / self.slide_width))

                pres_hash = hashlib.md5(pres_name.encode()).hexdigest()[:8]

                # Export current image
                if 1 <= current_idx <= pres.Slides.Count:
                    try:
                        curr_path = os.path.join(self.temp_dir, f"ppt_{pres_hash}_{current_idx}.jpg")
                        if not os.path.exists(curr_path):
                            slide = pres.Slides(current_idx)
                            slide.Export(curr_path, "JPG", scale_w, scale_h)
                        with open(curr_path, "rb") as f:
                            state["current_image"] = "data:image/jpeg;base64," + base64.b64encode(f.read()).decode('utf-8')
                    except Exception as e:
                        print("Error getting current image:", e)
                
                # Export next image
                if 1 <= current_idx + 1 <= pres.Slides.Count:
                    try:
                        next_path = os.path.join(self.temp_dir, f"ppt_{pres_hash}_{current_idx+1}.jpg")
                        if not os.path.exists(next_path):
                            slide = pres.Slides(current_idx+1)
                            slide.Export(next_path, "JPG", scale_w, scale_h)
                        with open(next_path, "rb") as f:
                            state["next_image"] = "data:image/jpeg;base64," + base64.b64encode(f.read()).decode('utf-8')
                    except Exception as e:
                        print("Error getting next image:", e)

                self.last_pres_name = pres_name
                self.last_slide_idx = current_idx
                self.last_state = state

        except Exception as e:
            print("Error getting state:", e)
            
        return state

    def get_slides_info(self):
        app = self._get_app()
        slides = []
        if app:
            try:
                if app.Presentations.Count > 0:
                    pres = app.ActivePresentation
                    for i in range(1, pres.Slides.Count + 1):
                        slide = pres.Slides(i)
                        title = f"Slide {i}"
                        if slide.Shapes.HasTitle:
                            if slide.Shapes.Title.TextFrame.HasText:
                                title = slide.Shapes.Title.TextFrame.TextRange.Text
                        slides.append({"index": i, "title": title})
            except Exception as e:
                print("Error getting slides info:", e)
        return slides

    def get_slide_thumb(self, index):
        app = self._get_app()
        if app:
            try:
                if app.Presentations.Count > 0:
                    pres = app.ActivePresentation
                    if 1 <= index <= pres.Slides.Count:
                        pres_hash = hashlib.md5(pres.Name.encode()).hexdigest()[:8]
                        path = os.path.join(self.temp_dir, f"ppt_{pres_hash}_thumb_{index}.jpg")
                        if not os.path.exists(path):
                            w = pres.PageSetup.SlideWidth
                            h = pres.PageSetup.SlideHeight
                            pres.Slides(index).Export(path, "JPG", 320, int(320 * (h/w)))
                        return path
            except Exception as e:
                print("Error exporting thumb:", e)
        return None

    def cleanup_cache(self):
        try:
            pattern = os.path.join(self.temp_dir, "ppt_*.jpg")
            for f in glob.glob(pattern):
                try:
                    os.remove(f)
                except Exception:
                    pass
        except Exception:
            pass
