import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re
from datetime import timedelta
from pathlib import Path


class TextToSRTConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Text to Subtitle Converter")
        self.root.geometry("500x520")
        self.root.resizable(False, False)

        self.input_file = None
        self.input_folder = None

        style = ttk.Style()
        style.theme_use('clam')

        ttk.Label(root, text="Text to Subtitle Converter",
                  font=("Arial", 14, "bold")).grid(
            row=0, column=0, columnspan=3, pady=10)

        ttk.Label(root, text="Single File:").grid(
            row=1, column=0, sticky="w", padx=10, pady=5)
        self.input_label = ttk.Label(
            root, text="No file selected", foreground="gray")
        self.input_label.grid(row=1, column=1, sticky="w", padx=5)
        ttk.Button(root, text="Browse",
                   command=self.select_input_file).grid(
            row=1, column=2, padx=5, pady=5)

        ttk.Label(root, text="Folder (Batch):").grid(
            row=2, column=0, sticky="w", padx=10, pady=5)
        self.folder_label = ttk.Label(
            root, text="No folder selected", foreground="gray")
        self.folder_label.grid(row=2, column=1, sticky="w", padx=5)
        ttk.Button(root, text="Browse",
                   command=self.select_input_folder).grid(
            row=2, column=2, padx=5, pady=5)

        ttk.Label(root, text="Output Format:").grid(
            row=3, column=0, sticky="w", padx=10, pady=5)
        self.format_var = tk.StringVar(value="VTT (WebVTT)")
        ttk.Combobox(root, textvariable=self.format_var,
                     values=["VTT (WebVTT)", "SRT"],
                     state="readonly", width=20).grid(
            row=3, column=1, columnspan=2, sticky="w", padx=5)

        ttk.Label(root, text="Max Words Per Caption:").grid(
            row=4, column=0, sticky="w", padx=10, pady=5)
        self.max_words_var = tk.IntVar(value=15)
        ttk.Spinbox(root, from_=5, to=50,
                    textvariable=self.max_words_var,
                    width=10, increment=1).grid(
            row=4, column=1, sticky="w", padx=5)

        ttk.Label(root, text="Timing Mode:").grid(
            row=5, column=0, sticky="w", padx=10, pady=5)
        self.timing_mode_var = tk.StringVar(value="Auto (use doc duration)")
        ttk.Combobox(root, textvariable=self.timing_mode_var,
                     values=["Auto (use doc duration)", "WPM (Smart)"],
                     state="readonly", width=22).grid(
            row=5, column=1, columnspan=2, sticky="w", padx=5)

        ttk.Label(root, text="Words Per Minute:").grid(
            row=6, column=0, sticky="w", padx=10, pady=5)
        self.wpm_var = tk.IntVar(value=140)
        ttk.Spinbox(root, from_=80, to=220,
                    textvariable=self.wpm_var,
                    width=10, increment=10).grid(
            row=6, column=1, sticky="w", padx=5)

        ttk.Label(root, text="Min Duration (sec):").grid(
            row=7, column=0, sticky="w", padx=10, pady=5)
        self.min_dur_var = tk.DoubleVar(value=1.5)
        ttk.Spinbox(root, from_=0.5, to=10.0,
                    textvariable=self.min_dur_var,
                    width=10, increment=0.5).grid(
            row=7, column=1, sticky="w", padx=5)

        ttk.Label(root, text="Max Duration (sec):").grid(
            row=8, column=0, sticky="w", padx=10, pady=5)
        self.max_dur_var = tk.DoubleVar(value=8.0)
        ttk.Spinbox(root, from_=1.0, to=30.0,
                    textvariable=self.max_dur_var,
                    width=10, increment=0.5).grid(
            row=8, column=1, sticky="w", padx=5)

        ttk.Button(root, text="Convert to Subtitle File",
                   command=self.convert).grid(
            row=9, column=0, columnspan=3, pady=20)

        ttk.Label(
            root,
            text=(
                "Select a single file OR a folder for batch processing.\n"
                "Batch mode creates a 'subtitles_output' folder next to\n"
                "the selected folder containing all converted files."
            ),
            font=("Arial", 9), foreground="gray",
            wraplength=450, justify="left"
        ).grid(row=10, column=0, columnspan=3, padx=10, pady=5, sticky="w")

    def select_input_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Text or Word File",
            filetypes=[
                ("Supported Files", "*.txt *.docx"),
                ("Text Files", "*.txt"),
                ("Word Documents", "*.docx"),
                ("All Files", "*.*"),
            ],
        )
        if file_path:
            self.input_file = file_path
            self.input_folder = None
            self.input_label.config(
                text=os.path.basename(file_path), foreground="black")
            self.folder_label.config(
                text="No folder selected", foreground="gray")

    def select_input_folder(self):
        folder_path = filedialog.askdirectory(
            title="Select Folder Containing DOCX Files"
        )
        if folder_path:
            self.input_folder = folder_path
            self.input_file = None
            self.folder_label.config(
                text=os.path.basename(folder_path), foreground="black")
            self.input_label.config(
                text="No file selected", foreground="gray")

    def time_to_format(self, seconds, format_type="vtt"):
        td = timedelta(seconds=seconds)
        hours, remainder = divmod(int(td.total_seconds()), 3600)
        minutes, seconds_part = divmod(remainder, 60)
        milliseconds = int((td.total_seconds() % 1) * 1000)
        separator = "." if format_type == "vtt" else ","
        return (f"{hours:02d}:{minutes:02d}:{seconds_part:02d}"
                f"{separator}{milliseconds:03d}")

    def extract_total_duration(self, paragraphs: list) -> float:
        """
        Look for duration pattern in document header.
        Matches formats like:
          Thu, Apr 16, 2026 3:45PM • 0:45
          Thu, Apr 16, 2026 3:45PM • 1:23:45
        Adds a 2 second buffer to catch trailing words.
        Returns total seconds as float or None if not found.
        """
        duration_pattern = re.compile(
            r'•\s*(\d+):(\d{2})(?::(\d{2}))?'
        )
        for line in paragraphs:
            match = duration_pattern.search(line)
            if match:
                part1 = int(match.group(1))
                part2 = int(match.group(2))
                part3 = match.group(3)
                if part3 is not None:
                    total = part1 * 3600 + part2 * 60 + int(part3)
                else:
                    total = part1 * 60 + part2
                total += 2.0
                return float(total)
        return None

    def split_into_captions(self, text: str) -> list:
        """
        Stage 1: Split continuous text into caption sized chunks.
        1. Split at sentence boundaries (. ! ?)
        2. Split long sentences at commas and semicolons
        3. Force split at max word count
        """
        max_words = self.max_words_var.get()
        raw_sentences = re.split(r'(?<=[.!?])\s+', text.strip())
        raw_sentences = [s.strip() for s in raw_sentences if s.strip()]
        captions = []
        for sentence in raw_sentences:
            words = sentence.split()
            if len(words) <= max_words:
                captions.append(sentence)
            else:
                sub_parts = re.split(r'(?<=[,;])\s+', sentence)
                current_chunk = []
                for part in sub_parts:
                    part_words = part.split()
                    if len(current_chunk) + len(part_words) <= max_words:
                        current_chunk.extend(part_words)
                    else:
                        if current_chunk:
                            captions.append(" ".join(current_chunk))
                        if len(part_words) > max_words:
                            for i in range(0, len(part_words), max_words):
                                chunk = " ".join(
                                    part_words[i:i + max_words])
                                if chunk:
                                    captions.append(chunk)
                            current_chunk = []
                        else:
                            current_chunk = part_words
                if current_chunk:
                    captions.append(" ".join(current_chunk))
        return [c for c in captions if c.strip()]

    def calculate_durations_proportional(
            self, captions: list, total_seconds: float) -> list:
        """
        Distribute total_seconds proportionally by word count.
        Only applies minimum guardrail — no maximum cap so the
        total always sums correctly to total_seconds.
        Normalises after applying minimum to guarantee exact total.
        """
        word_counts = [len(c.split()) for c in captions]
        total_words = sum(word_counts)
        min_dur = self.min_dur_var.get()
        if total_words == 0:
            return [min_dur] * len(captions)
        durations = []
        for wc in word_counts:
            dur = (wc / total_words) * total_seconds
            dur = max(dur, min_dur)
            durations.append(dur)
        current_total = sum(durations)
        if current_total > 0:
            scale = total_seconds / current_total
            durations = [d * scale for d in durations]
        return durations

    def calculate_duration_wpm(self, text: str) -> float:
        """Calculate duration from WPM for a single caption."""
        word_count = len(text.split())
        wpm = self.wpm_var.get()
        duration = (word_count / wpm) * 60.0
        duration = max(duration, self.min_dur_var.get())
        duration = min(duration, self.max_dur_var.get())
        return duration

    def process_single_file(self, input_path: str,
                             output_path: str) -> tuple:
        """
        Process one docx or txt file and write subtitle output.
        Returns (success: bool, message: str)
        """
        try:
            from docx import Document
        except ImportError:
            return False, (
                "python-docx is not installed.\n"
                "Run: pip install python-docx"
            )

        file_ext = os.path.splitext(input_path)[1].lower()

        if file_ext == ".docx":
            try:
                doc = Document(input_path)
            except Exception as e:
                return False, (
                    f"Could not open "
                    f"{os.path.basename(input_path)}: {e}")

            all_paragraphs = [
                para.text.strip() for para in doc.paragraphs
            ]
            total_duration = self.extract_total_duration(all_paragraphs)

            content_paragraphs = []
            found_start = False
            for text in all_paragraphs:
                if not found_start:
                    if text == "00:00":
                        found_start = True
                        continue
                    continue
                if text:
                    content_paragraphs.append(text)

            if not found_start:
                return False, (
                    f"No '00:00' marker found in "
                    f"{os.path.basename(input_path)} — skipped."
                )

            content = " ".join(content_paragraphs)

        else:
            content = None
            for encoding in ["utf-8", "utf-16", "utf-8-sig",
                              "latin-1", "cp1252"]:
                try:
                    with open(input_path, "r", encoding=encoding) as f:
                        content = f.read()
                    break
                except (UnicodeDecodeError, UnicodeError):
                    continue
            if content is None:
                with open(input_path, "r", encoding="utf-8",
                          errors="replace") as f:
                    content = f.read()
            content = " ".join(content.split())
            total_duration = None

        subtitle_texts = self.split_into_captions(content)
        if not subtitle_texts:
            return False, (
                f"No caption content found in "
                f"{os.path.basename(input_path)} — skipped."
            )

        timing_mode = self.timing_mode_var.get()
        if (timing_mode == "Auto (use doc duration)"
                and total_duration is not None):
            durations = self.calculate_durations_proportional(
                subtitle_texts, total_duration)
        else:
            durations = [
                self.calculate_duration_wpm(t) for t in subtitle_texts
            ]

        format_display = self.format_var.get()
        format_type = "vtt" if format_display == "VTT (WebVTT)" else "srt"

        content_lines = []
        if format_type == "vtt":
            content_lines.append("WEBVTT\n")

        current_time = 0.0
        for index, (subtitle_text, duration) in enumerate(
                zip(subtitle_texts, durations), 1):
            start_time = current_time
            end_time = current_time + duration
            start_str = self.time_to_format(start_time, format_type)
            end_str = self.time_to_format(end_time, format_type)
            if format_type == "srt":
                content_lines.append(str(index))
            content_lines.append(f"{start_str} --> {end_str}")
            content_lines.append(subtitle_text)
            content_lines.append("")
            current_time = end_time

        with open(output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(content_lines))

        return True, (
            f"{os.path.basename(input_path)} → "
            f"{len(subtitle_texts)} captions"
        )

    def convert(self):
        if not self.input_file and not self.input_folder:
            messagebox.showerror(
                "Error", "Please select an input file or folder")
            return

        format_display = self.format_var.get()
        format_type = "vtt" if format_display == "VTT (WebVTT)" else "srt"
        file_extension = format_type

        try:
            # Single file mode
            if self.input_file:
                input_dir = os.path.dirname(self.input_file)
                input_name = os.path.splitext(
                    os.path.basename(self.input_file))[0]
                output_path = os.path.join(
                    input_dir, f"{input_name}.{file_extension}")
                success, message = self.process_single_file(
                    self.input_file, output_path)
                if success:
                    messagebox.showinfo(
                        "Success",
                        f"File created successfully!\n{message}\n\n"
                        f"{output_path}")
                else:
                    messagebox.showerror("Error", message)

            # Batch folder mode
            else:
                docx_files = list(
                    Path(self.input_folder).rglob("*.docx"))
                if not docx_files:
                    messagebox.showerror(
                        "No Files Found",
                        "No .docx files found in the selected folder.")
                    return

                parent_dir = os.path.dirname(self.input_folder)
                folder_name = os.path.basename(self.input_folder)
                output_dir = os.path.join(
                    parent_dir, f"{folder_name}_subtitles_output")
                os.makedirs(output_dir, exist_ok=True)

                success_count = 0
                failed = []
                results = []

                for docx_path in docx_files:
                    output_name = docx_path.stem + f".{file_extension}"
                    output_path = os.path.join(output_dir, output_name)
                    success, message = self.process_single_file(
                        str(docx_path), output_path)
                    if success:
                        success_count += 1
                        results.append(f"✓ {message}")
                    else:
                        failed.append(f"✗ {message}")

                summary = (
                    f"Batch conversion complete!\n\n"
                    f"Converted: {success_count} of "
                    f"{len(docx_files)} files\n"
                    f"Output folder:\n{output_dir}\n\n"
                )
                if results:
                    summary += "Successful:\n" + "\n".join(results)
                if failed:
                    summary += "\n\nFailed:\n" + "\n".join(failed)

                messagebox.showinfo("Batch Complete", summary)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = TextToSRTConverter(root)
    root.mainloop()