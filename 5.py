    def choose_txt_files(self) -> None:
        initial = self.folder_var.get() or str(Path.cwd())
        selected = filedialog.askopenfilenames(
            title="Lxy/WST TXT 파일 직접 선택",
            initialdir=initial,
            filetypes=[("Lxy/WST text", "*.txt"), ("All files", "*.*")],
        )
        if not selected:
            return
        paths = [Path(path) for path in selected]
        folders = {path.parent for path in paths}
        if len(folders) != 1:
            messagebox.showerror("TXT 선택 오류", "Lxy/WST TXT는 같은 폴더 안의 파일로 선택해 주세요.")
            return
        folder = paths[0].parent
        lxy_files = [path for path in paths if path.name.lower().endswith("_lxy.txt")]
        wst_files = [path for path in paths if path.name.lower().endswith("_wst.txt")]
        if not lxy_files or not wst_files:
            messagebox.showerror(
                "TXT 선택 오류",
                "파일명은 *_Lxy.txt 와 *_WST.txt 형식이어야 합니다.\n예: 3M#1_Lxy.txt / 3M#1_WST.txt",
            )
            return
        stems = {path.name[:-8].casefold() for path in lxy_files}
        wst_stems = {path.name[:-8].casefold() for path in wst_files}
        matched = stems & wst_stems
        if not matched:
            messagebox.showerror(
                "TXT 선택 오류",
                "선택한 TXT에서 Lxy/WST 짝을 찾지 못했습니다.\n같은 이름의 *_Lxy.txt / *_WST.txt를 함께 선택해 주세요.",
            )
            return
        self.manual_pcf_files = []
        self.folder_var.set(str(folder))
        self.reference_var.set(str(suggest_reference_excel(folder)))
        output = automatic_output_path(folder)
        self.output_var.set(str(output))
        self.output_name_var.set(output.name)
        self.force_pcf_var.set(False)
        self.refresh_cards()
        self.refresh_sample_options()
        self.status_var.set(f"Lxy/WST TXT {len(matched)}쌍을 선택했습니다. TXT로 분석합니다.")

    def choose_pcf_files(self) -> None:
        initial = self.folder_var.get() or str(Path.cwd())
        selected = filedialog.askopenfilenames(
            title="PCF/PCT 파일 직접 선택",
            initialdir=initial,
            filetypes=[("LMK PCF/PCT", "*.pcf *.pct"), ("All files", "*.*")],
        )
        if not selected:
            return
        self.manual_pcf_files = [Path(path) for path in selected]
        first_parent = self.manual_pcf_files[0].parent
        self.folder_var.set(str(first_parent))
        self.reference_var.set(str(suggest_reference_excel(first_parent)))
        output = automatic_output_path(first_parent)
        self.output_var.set(str(output))
        self.output_name_var.set(output.name)
        self.force_pcf_var.set(True)
        self.refresh_cards()
        self.refresh_sample_options()
        self.status_var.set(f"PCF/PCT 파일 {len(self.manual_pcf_files)}개를 직접 선택했습니다.")

    def clear_pcf_files(self) -> None:
        self.manual_pcf_files = []
        self.refresh_cards()
        self.refresh_sample_options()
        self.status_var.set("수동 선택 PCF를 해제했습니다. Data 폴더 자동 검색을 사용합니다.")

    def stop_work(self) -> None:
        self.cancel_event.set()
        self.status_var.set("중지 요청됨: 현재 시료 처리 후 멈춥니다.")
        if self.stop_button:
            self.stop_button.configure(state="disabled")

    def update_result_preview(self) -> None:
        if not self.result_text:
            return
        selected_id = parse_sample_selection(self.result_sample_var.get())
        sample = next((item for item in self.last_results if item.sample_id == selected_id), None)
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", tk.END)
        if sample is None:
            self.result_text.insert("1.0", "분석 완료 후 선택한 시료의 15 x 35 MPCD 결과가 표시됩니다.")
            self.result_text.configure(state="disabled")
            return
        header = "Row " + " ".join(f"{col:>8}" for col in range(1, sample.cols + 1))
        lines = [f"{sample_title(sample)}  ({sample.source})", header]
        for row_index, row in enumerate(sample.values, 1):
            values = ["        " if value is None else f"{value:8.2f}" for value in row]
            lines.append(f"{row_index:>3} " + " ".join(values))
        self.result_text.insert("1.0", "\n".join(lines))
        self.result_text.configure(state="disabled")

    def set_layout(self, value: str) -> None:
        self.layout_var.set(value)
        names = {"excel_style": "정리 Excel형", "vertical": "전체 아래로"}
        self.status_var.set(f"보기 방식: {names.get(value, value)}")

    def open_path(self, path: Path) -> None:
        try:
            if not path.exists():
                messagebox.showinfo("확인", f"아직 존재하지 않습니다:\n{path}")
                return
            os.startfile(path)
        except Exception as exc:
            messagebox.showerror("열기 실패", str(exc))

    def choose_folder(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.folder_var.get())
        if selected:
            folder = Path(selected)
            self.manual_pcf_files = []
            self.folder_var.set(str(folder))
            self.reference_var.set(str(suggest_reference_excel(folder)))
            output = automatic_output_path(folder)
            self.output_var.set(str(output))
            self.output_name_var.set(output.name)
            self.refresh_cards()
            self.refresh_sample_options()
            self.status_var.set("데이터 폴더를 선택했습니다.")

    def choose_labsoft_path(self) -> None:
        initial = self.labsoft_path_var.get() or default_labsoft_path() or str(Path.home())
        selected = filedialog.askopenfilename(
            title="LMK LabSoft 선택",
            initialdir=str(Path(initial).parent if initial else Path.home()),
            filetypes=[("LabSoft 실행/바로가기", "*.exe *.lnk"), ("All files", "*.*")],
        )
        if selected:
            self.labsoft_path_var.set(selected)
            self.settings["labsoft_path"] = selected
            save_app_settings(self.settings)
            self.refresh_cards()
            self.status_var.set("LMK LabSoft 경로를 저장했습니다.")

    def choose_reference(self) -> None:
        selected = filedialog.askopenfilename(filetypes=[("Excel workbook", "*.xlsx")], initialdir=self.folder_var.get())
        if selected:
            self.reference_var.set(selected)
            self.refresh_cards()
            self.refresh_sample_options()
            self.status_var.set("기준 Excel을 선택했습니다.")

    def choose_output(self) -> None:
        current = Path(self.output_var.get())
        selected = filedialog.askdirectory(initialdir=str(current.parent))
        if selected:
            file_name = self.output_name_var.get().strip() or current.name
            self.output_var.set(str(Path(selected) / file_name))
            self.refresh_cards()
            self.status_var.set("결과 폴더를 선택했습니다.")

    def reset_form(self) -> None:
        cwd = Path.cwd()
        self.panel_roi_var.set("0.98")
        self.manual_pcf_files = []
        self.folder_var.set(str(cwd))
        self.labsoft_path_var.set(self.settings.get("labsoft_path", default_labsoft_path()))
        self.force_pcf_var.set(False)
        self.reference_var.set(str(suggest_reference_excel(cwd)))
        output = automatic_output_path(cwd)
        self.output_var.set(str(output))
        self.output_name_var.set(output.name)
        self.layout_var.set("excel_style")
        self.last_output = None
        self.last_csv = None
        self.last_results = []
        self.result_sample_var.set("")
        if self.progress_bar:
            self.progress_bar.configure(value=0)
        if self.result_title_label:
            self.result_title_label.configure(text="15 x 35 MPCD 결과")
        if self.result_detail_label:
            self.result_detail_label.configure(text="분석을 실행하면 Excel / CSV 결과를 여기서 바로 열 수 있습니다.")
        if self.result_sample_combo:
            self.result_sample_combo.configure(values=())
        self.update_result_preview()
        self.refresh_cards()
        self.refresh_sample_options()
        self.status_var.set("초기화 완료")

    def write_error_log(self, output: Path, message: str) -> Optional[Path]:
        try:
            log_path = output.parent / "MPCD_error_log.txt"
            log_path.write_text(message, encoding="utf-8")
            return log_path
        except Exception:
            return None

    def show_error(self, brief: str, detail: str, log_path: Optional[Path]) -> None:
        log_text = f"\n\n상세 로그: {log_path}" if log_path else ""
        messagebox.showerror("오류", f"{brief}{log_text}\n\n{detail}")

    def open_last_output(self) -> None:
        if not self.last_output:
            messagebox.showinfo("결과 없음", "먼저 Excel을 생성하세요.")
            return
        self.open_path(self.last_output)

    def open_last_csv(self) -> None:
        if not self.last_csv:
            messagebox.showinfo("결과 없음", "먼저 결과를 생성하세요.")
            return
        self.open_path(self.last_csv)

    def open_output_folder(self) -> None:
        target = self.last_output.parent if self.last_output else Path(self.output_var.get()).parent
        self.open_path(target)

    def start(self) -> None:
        folder = Path(self.folder_var.get())
        reference = Path(self.reference_var.get())
        output = self.resolve_output_path()
        layout = self.layout_var.get()
        sample_selection = self.get_sample_selection()
        time_groups: List[str] = []
        labsoft_path = self.labsoft_path_var.get().strip()
        force_pcf_for_run = self.force_pcf_var.get()
        if force_pcf_for_run:
            existing_ids = existing_txt_sample_ids(folder, sample_selection, self.manual_pcf_files)
            if existing_ids:
                sample_text = ", ".join(str(sample_id) for sample_id in existing_ids[:12])
                if len(existing_ids) > 12:
                    sample_text += f" 외 {len(existing_ids) - 12}개"
                answer = messagebox.askyesnocancel(
                    "TXT 재생성 확인",
                    "이미 생성된 Lxy/WST txt가 있습니다.\n\n"
                    f"대상 Sample: {sample_text}\n\n"
                    "PCF에서 다시 추출해 txt를 새로 생성할까요?\n"
                    "예: PCF에서 다시 추출\n"
                    "아니오: 기존 txt로 분석\n"
                    "취소: 분석 취소",
                )
                if answer is None:
                    self.status_var.set("분석 취소")
                    return
                force_pcf_for_run = bool(answer)
        os.environ["MPCD_FORCE_PCF"] = "1" if force_pcf_for_run else "0"
        source_mode = "pcf" if force_pcf_for_run else "raw_txt"
        self.settings["force_pcf"] = "1" if self.force_pcf_var.get() else "0"
        if labsoft_path:
            os.environ["MPCD_LABSOFT_PATH"] = labsoft_path
            self.settings["labsoft_path"] = labsoft_path
        try:
            save_app_settings(self.settings)
        except Exception:
            pass
        try:
            panel_roi = float(self.panel_roi_var.get().strip())
        except ValueError:
            messagebox.showerror("오류", "Panel ROI는 숫자로 입력해 주세요. 예: 0.98")
            return
        if self.create_button:
            self.create_button.configure(state="disabled", text="분석 중...")
        if self.stop_button:
            self.stop_button.configure(state="normal")
        self.cancel_event.clear()
        if self.progress_bar:
            self.progress_bar.configure(value=15)
        if self.result_meta_label:
            self.result_meta_label.configure(text=self.result_meta_text())

        def worker() -> None:
            try:
                def progress(text: str) -> None:
                    if self.cancel_event.is_set():
                        raise OperationCancelled("사용자가 분석을 중지했습니다.")
                    self.set_status(text)

                results = run_export(
                    folder,
                    output,
                    layout,
                    source_mode,
                    reference,
                    time_groups,
                    sample_selection,
                    panel_roi,
                    progress,
                    self.manual_pcf_files,
                )
                self.last_output = output
                self.last_csv = output.with_suffix(".csv")
                self.last_results = list(results)
                self.after(0, self.show_result)
            except OperationCancelled as exc:
                self.set_status(str(exc))
                if self.progress_bar:
                    self.after(0, lambda: self.progress_bar.configure(value=0))
            except Exception as exc:
                brief, user_message = user_friendly_error(exc)
                detail = traceback.format_exc()
                log_message = f"{brief}\n\n{user_message}\n\n--- technical detail ---\n{detail}"
                log_path = self.write_error_log(output, log_message)
                self.set_status(f"오류: {brief}")
                if self.progress_bar:
                    self.after(0, lambda: self.progress_bar.configure(value=0))
                self.after(0, lambda brief=brief, detail=user_message, log_path=log_path: self.show_error(brief, detail, log_path))
            finally:
                if self.create_button:
                    self.after(0, lambda: self.create_button.configure(state="normal", text="분석"))
                if self.stop_button:
                    self.after(0, lambda: self.stop_button.configure(state="disabled"))

        threading.Thread(target=worker, daemon=True).start()

    def set_status(self, text: str) -> None:
        def update() -> None:
            self.status_var.set(text)
            if self.progress_bar:
                if "읽" in text:
                    self.progress_bar.configure(value=35)
                elif "저장" in text:
                    self.progress_bar.configure(value=75)

        self.after(0, update)

    def show_result(self) -> None:
        if self.result_title_label:
            self.result_title_label.configure(text="15 x 35 MPCD 결과")
        if self.result_meta_label:
            self.result_meta_label.configure(text=self.result_meta_text())
        if self.result_detail_label and self.last_output:
            csv_name = self.last_csv.name if self.last_csv else self.last_output.with_suffix(".csv").name
            self.result_detail_label.configure(
                text=f"분석 완료: {self.last_output.name} / {csv_name} 파일을 바로 열 수 있습니다."
            )
        if self.result_sample_combo:
            values = [f"{sample_title(sample)} (Sample {sample.sample_id})" for sample in self.last_results]
            self.result_sample_combo.configure(values=values)
            if values:
                self.result_sample_var.set(values[0])
        self.update_result_preview()
        if self.progress_bar:
            self.progress_bar.configure(value=100)
        self.status_var.set("완료")

    def resolve_output_path(self) -> Path:
        output = available_output_path(automatic_output_path(Path(self.folder_var.get())))
        self.output_var.set(str(output))
        self.output_name_var.set(output.name)
        self.refresh_cards()
        return output


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create MPCD Excel output matched to 정리.xlsx.")
    parser.add_argument("--folder", default=os.getcwd(), help="data folder")
    parser.add_argument("--reference-excel", default=None, help="reference 정리.xlsx path")
    parser.add_argument("--output", default=None, help="output .xlsx path")
    parser.add_argument("--layout", choices=("excel_style", "vertical", "by_time"), default="excel_style")
    parser.add_argument("--sample", default="전체", help="sample selection, e.g. '전체' or 'Sample 3'")
    parser.add_argument("--time-groups", default="", help="comma-separated group labels, e.g. '0hr,240hr'")
    parser.add_argument("--roi", type=float, default=0.98, help="Panel ROI ratio, default 0.98")
    parser.add_argument("--labsoft-path", default="", help="LMK LabSoft .exe or .lnk path for PCF auto export")
    parser.add_argument("--pcf", action="append", default=[], help="manually selected PCF/PCT path; can be repeated")
    parser.add_argument("--source", choices=("pcf", "txt"), default="txt", help="pcf: export from PCF first, txt: use existing txt first")
    parser.add_argument("--no-gui", action="store_true", help="run without GUI")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    folder = Path(args.folder).resolve()
    reference = Path(args.reference_excel).resolve() if args.reference_excel else None
    output = Path(args.output).resolve() if args.output else available_output_path(automatic_output_path(folder))
    groups = [part.strip() for part in args.time_groups.split(",") if part.strip()]
    os.environ["MPCD_FORCE_PCF"] = "1" if args.source == "pcf" else "0"
    if args.labsoft_path:
        os.environ["MPCD_LABSOFT_PATH"] = args.labsoft_path
    if args.no_gui:
        source_mode = "pcf" if args.source == "pcf" else "raw_txt"
        manual_pcf_files = [Path(path).resolve() for path in args.pcf]
        run_export(folder, output, args.layout, source_mode, reference, groups, args.sample, args.roi, print, manual_pcf_files)
        print(f"Saved: {output}")
        sys.stdout.flush()
        sys.stderr.flush()
        if getattr(sys, "frozen", False):
            os._exit(0)
        return 0
    App().mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
