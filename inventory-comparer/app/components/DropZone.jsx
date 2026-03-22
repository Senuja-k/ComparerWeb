"use client";

import { useRef, useState } from "react";

export default function DropZone({
  title,
  icon,
  description,
  files,
  onFilesChange,
  multiple = true,
  accept = ".xlsx,.xls",
  accentColor = "#4a6dff",
}) {
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef(null);

  const acceptPattern = accept
    .split(",")
    .map((ext) => ext.trim().replace(".", ""))
    .join("|");
  const regex = new RegExp(`\\.(${acceptPattern})$`, "i");

  function handleDrop(e) {
    e.preventDefault();
    setDragOver(false);
    const dropped = [...e.dataTransfer.files].filter((f) => regex.test(f.name));
    if (dropped.length === 0) return;

    if (multiple) {
      const existing = new Set(files.map((f) => f.name));
      const newFiles = dropped.filter((f) => !existing.has(f.name));
      onFilesChange([...files, ...newFiles]);
    } else {
      onFilesChange([dropped[0]]);
    }
  }

  function handleInputChange(e) {
    const selected = e.target.files ? [...e.target.files].filter((f) => regex.test(f.name)) : [];
    if (selected.length === 0) return;

    if (multiple) {
      const existing = new Set(files.map((f) => f.name));
      const newFiles = selected.filter((f) => !existing.has(f.name));
      onFilesChange([...files, ...newFiles]);
    } else {
      onFilesChange([selected[0]]);
    }
    if (inputRef.current) inputRef.current.value = "";
  }

  function removeFile(name) {
    onFilesChange(files.filter((f) => f.name !== name));
  }

  return (
    <div
      className={`flex-1 bg-white border-2 border-dashed rounded-xl p-6 text-center min-h-[250px] transition-all cursor-pointer ${
        dragOver ? "border-solid" : ""
      }`}
      style={{ borderColor: dragOver ? accentColor : "#dee2e6" }}
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={handleDrop}
      onClick={() => inputRef.current?.click()}
    >
      <input
        type="file"
        ref={inputRef}
        className="hidden"
        accept={accept}
        multiple={multiple}
        onChange={handleInputChange}
      />
      <h2 className="text-xl font-semibold mb-2">
        {icon} {title}
      </h2>
      <p className="text-sm text-[#787878]">{description}</p>
      <div className="mt-4 text-left max-h-[150px] overflow-y-auto">
        {files.map((f) => (
          <div
            key={f.name}
            className="bg-[#f5f9ff] rounded-md px-3 py-1.5 mb-1.5 text-sm flex justify-between items-center"
          >
            <span>{f.name.replace(/\.(xlsx|xls|csv)$/i, "")}</span>
            <span
              className="cursor-pointer text-[#ff6b6b] font-bold ml-2"
              onClick={(e) => {
                e.stopPropagation();
                removeFile(f.name);
              }}
            >
              ✖
            </span>
          </div>
        ))}
      </div>
    </div>
  );
}
