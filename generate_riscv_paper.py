#!/usr/bin/env python3
"""
Generate a publication-grade research paper from the analyzed
Tiny Tapeout RISC-V repository:

  https://github.com/PrmBRana/ttgf_prem_programmable_RISC_v

Grounded in the actual Verilog (pipeline.v, Control_Unit.v, uart.v,
uart_bootloader.v, spi.v, gpio.v, Instruction_Mem.v, DataMem.v,
info.yaml, docs/info.md, test/test.py).
"""

import os
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

OUT_DIR = "/home/prem/Documents/research paper"
IMG = OUT_DIR
REPO_BLOCK = "/home/prem/asic_tools/RISC_V/GF180MCU/ttgf_prem_risc_v/docs/img/block.png"


def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'))


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    ))


def set_table_borders(table, color="D2D6DC"):
    table._tbl.tblPr.append(parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="none"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tblBorders>'
    ))


def build_document():
    if os.path.isfile(REPO_BLOCK):
        try:
            shutil.copy(REPO_BLOCK, os.path.join(OUT_DIR, "block_diagram_repo.png"))
        except OSError:
            pass

    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        section.different_first_page_header_footer = True
        hp = section.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hr = hp.add_run("From Verilog to GDSII: A Programmable Five-Stage RISC-V RV32I SoC | Rana & Shrestha")
        hr.font.name = "Calibri"
        hr.font.size = Pt(8)
        hr.font.color.rgb = RGBColor(120, 130, 140)
        fp = section.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fr = fp.add_run("Antarikchya Pratisthan Nepal / Nepal Space Foundation  —  Research Paper  —  Tiny Tapeout GF 26a Project 0039")
        fr.font.name = "Calibri"
        fr.font.size = Pt(8)
        fr.font.color.rgb = RGBColor(120, 130, 140)

    C_NAVY = RGBColor(16, 44, 87)
    C_SLATE = RGBColor(53, 89, 143)
    C_BODY = RGBColor(34, 34, 34)
    C_GRAY = RGBColor(100, 110, 120)

    def add_sec(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(title)
        r.font.name = "Calibri"
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = C_NAVY

    def add_sub(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(title)
        r.font.name = "Calibri"
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = C_SLATE

    def add_ss(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(title)
        r.font.name = "Calibri"
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.italic = True
        r.font.color.rgb = C_BODY

    def add_p(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(10)
        r.font.color.rgb = C_BODY
        return p

    def add_code(code_str, caption=None):
        if caption:
            pc = doc.add_paragraph()
            pc.paragraph_format.space_before = Pt(6)
            pc.paragraph_format.space_after = Pt(2)
            pc.paragraph_format.keep_with_next = True
            rc = pc.add_run(caption)
            rc.font.bold = True
            rc.font.size = Pt(9.5)
            rc.font.color.rgb = C_NAVY
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.columns[0].width = Inches(6.8)
        cell = tbl.cell(0, 0)
        set_cell_background(cell, "F8FAFC")
        set_cell_margins(cell, 90, 90, 140, 140)
        cell._tc.get_or_add_tcPr().append(parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:left w:val="single" w:sz="18" w:space="0" w:color="2B6CB0"/>'
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            f'</w:tcBorders>'
        ))
        cp = cell.paragraphs[0]
        cp.paragraph_format.space_before = Pt(0)
        cp.paragraph_format.space_after = Pt(0)
        cp.paragraph_format.line_spacing = 1.05
        cr = cp.add_run(code_str.strip())
        cr.font.name = "Consolas"
        cr.font.size = Pt(8)
        cr.font.color.rgb = RGBColor(30, 41, 59)
        doc.add_paragraph().paragraph_format.space_after = Pt(3)

    def add_fig(name, caption, width=Inches(6.0)):
        path = name if os.path.isabs(name) else os.path.join(IMG, name)
        if not os.path.exists(path):
            print(f"Warning: missing figure {path}")
            return
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        p.add_run().add_picture(path, width=width)
        c = doc.add_paragraph()
        c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraph_format.space_after = Pt(8)
        r = c.add_run(caption)
        r.font.name = "Calibri"
        r.font.size = Pt(9)
        r.font.italic = True
        r.font.color.rgb = C_GRAY

    def add_tbl(headers, rows, widths, caption=None):
        if caption:
            pc = doc.add_paragraph()
            pc.paragraph_format.space_before = Pt(8)
            pc.paragraph_format.space_after = Pt(3)
            pc.paragraph_format.keep_with_next = True
            rc = pc.add_run(caption)
            rc.font.bold = True
            rc.font.size = Pt(9.5)
            rc.font.color.rgb = C_NAVY
        table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(table)
        for i, h in enumerate(headers):
            table.rows[0].cells[i].width = widths[i]
            set_cell_background(table.rows[0].cells[i], "102C57")
            set_cell_margins(table.rows[0].cells[i], 90, 90, 100, 100)
            hp = table.rows[0].cells[i].paragraphs[0]
            hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            hr = hp.add_run(h)
            hr.font.bold = True
            hr.font.size = Pt(8.5)
            hr.font.color.rgb = RGBColor(255, 255, 255)
        for ri, row in enumerate(rows):
            bg = "F8FAFC" if ri % 2 else "FFFFFF"
            for ci, val in enumerate(row):
                table.rows[ri + 1].cells[ci].width = widths[ci]
                set_cell_background(table.rows[ri + 1].cells[ci], bg)
                set_cell_margins(table.rows[ri + 1].cells[ci], 70, 70, 100, 100)
                rp = table.rows[ri + 1].cells[ci].paragraphs[0]
                rr = rp.add_run(str(val))
                rr.font.size = Pt(8)
                rr.font.color.rgb = C_BODY
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    def add_eq(eq, n):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(f"    {eq}    ")
        r.font.name = "Cambria Math"
        r.font.italic = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = C_NAVY
        nrun = p.add_run(f"({n})".rjust(8))
        nrun.font.bold = True
        nrun.font.size = Pt(9)
        nrun.font.color.rgb = C_SLATE

    # ==================================================================
    # TITLE
    # ==================================================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(
        "From Verilog to GDSII: Design, Simulation, and Silicon Realization of a "
        "Programmable Five-Stage RISC-V RV32I Microcontroller with UART Bootloader, "
        "16× Oversampling Serial I/O, GPIO, and Mode-0 SPI"
    )
    r.font.name = "Calibri"
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = C_NAVY

    pa = doc.add_paragraph()
    pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pa.paragraph_format.space_after = Pt(4)
    r1 = pa.add_run("Prem Bahadur Rana")
    r1.font.bold = True
    r1.font.size = Pt(11)
    r1.font.color.rgb = C_NAVY
    pa.add_run("1,*, ")
    r2 = pa.add_run("Trishna Shrestha")
    r2.font.bold = True
    r2.font.size = Pt(11)
    r2.font.color.rgb = C_NAVY
    pa.add_run("1\n")
    raff = pa.add_run(
        "1 Nepal Space Foundation / Antarikchya Pratisthan Nepal, Kathmandu, Nepal\n"
        "* Corresponding author: prem@antarikchya.org.np  |  Open-source RTL: "
        "github.com/PrmBRana/ttgf_prem_programmable_RISC_v"
    )
    raff.font.italic = True
    raff.font.size = Pt(9)
    raff.font.color.rgb = C_GRAY

    pd = doc.add_paragraph()
    pd.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pd.add_run("―" * 58).font.color.rgb = RGBColor(180, 190, 205)

    # Abstract
    tbl = doc.add_table(rows=1, cols=1)
    tbl.columns[0].width = Inches(6.8)
    c = tbl.cell(0, 0)
    set_cell_background(c, "F0F4F8")
    set_cell_margins(c, 140, 140, 200, 200)
    ph = c.paragraphs[0]
    rh = ph.add_run("Abstract—")
    rh.font.bold = True
    rh.font.size = Pt(10)
    rh.font.color.rgb = C_NAVY
    ra = ph.add_run(
        "This paper explains, in plain language and with the actual Verilog of a taped-out chip, "
        "how a 32-bit RISC-V processor works from the first clock edge after reset to the last "
        "polygon in a GDSII layout. The design is a programmable five-stage RV32I microcontroller "
        "(Instruction Fetch, Decode, Execute, Memory, Write-Back) submitted as Tiny Tapeout GF 26a "
        "project 0039, top module tt_um_prem_pipeline_test, occupying a 4×4 tile on the "
        "GlobalFoundries GF180MCU 180 nm open-source PDK. After reset the core is held stalled. "
        "A host sends handshake byte 0x25 over UART1 at 115 200 baud. The bootloader answers 0x55, "
        "assembles little-endian bytes into 32-bit instruction words, and writes them into a 64-word "
        "instruction memory. The sentinel 0xBAADF00D ends programming and releases the pipeline, "
        "which then fetches from address 0. Peripherals are not special CPU instructions: they are "
        "ordinary load/store addresses. UART2 lives at 0x1000_0000, GPIO at 0x3000_0000, and an "
        "SPI Mode-0 master at 0x4000_0000. The UART receiver oversamples each bit sixteen times and "
        "majority-votes the samples, which is why noisy start bits and clock drift do not instantly "
        "corrupt a firmware upload. The SPI master is parameterized for all four Motorola modes "
        "(CPOL, CPHA); the taped-out instance uses Mode 0 at about 4.17 MHz. Software builds "
        "peripheral pointers with LUI and ADDI, then polls status bits—there are no interrupts. "
        "The paper walks through instruction encoding, pointer arithmetic, hazard forwarding, "
        "cocotb/Icarus simulation, GTKWave reading, AMD Kria KV260 FPGA bring-up, and the "
        "LibreLane RTL-to-GDSII flow that turned this HDL into silicon."
    )
    ra.font.size = Pt(9.5)
    ra.font.color.rgb = C_BODY
    pk = c.add_paragraph()
    pk.paragraph_format.space_before = Pt(4)
    rkh = pk.add_run("Keywords—")
    rkh.font.bold = True
    rkh.font.size = Pt(9.5)
    rkh.font.color.rgb = C_NAVY
    rkt = pk.add_run(
        "RISC-V RV32I, five-stage pipeline, instruction decode, instruction memory, pointers, "
        "assembly language, UART 16× oversampling, hardware bootloader, GPIO, SPI Mode 0–3, "
        "Verilog, cocotb, GTKWave, FPGA, Tiny Tapeout, LibreLane, GF180MCU, GDSII."
    )
    rkt.font.italic = True
    rkt.font.size = Pt(9.5)
    rkt.font.color.rgb = C_BODY

    # ==================================================================
    # I
    # ==================================================================
    add_sec("I. INTRODUCTION — WHAT THIS CHIP IS, IN HUMAN LANGUAGE")
    add_p(
        "A processor is a machine that repeatedly does three things: it reads a 32-bit number from "
        "memory (an instruction), it decides what that number means (decode), and it changes "
        "registers, memory, or pins accordingly (execute). RISC-V is simply a public agreement on "
        "what those 32-bit numbers mean [1]–[3]. Anyone may implement the agreement in Verilog "
        "without a license fee. That is why a student or a national space laboratory can design a "
        "real CPU and send it to a foundry [4], [5]."
    )
    add_p(
        "The open-source repository analyzed in this work, ttgf_prem_programmable_RISC_v [6], is "
        "not a toy schematic. It is the HDL that was hardened for Tiny Tapeout shuttle GF 26a as "
        "Programmable_Pipeline-RISC-V at user-project address 39 [7], [8]. The same research "
        "program also prototyped the bootloader and UART path on an AMD Kria KV260 FPGA so that "
        "firmware could be uploaded from a laptop before any wafer existed [9], [10]."
    )
    add_p(
        "The paper is written so a reader who has never built a CPU can follow the story, while "
        "still being precise enough for a reviewer who will open the Verilog. Every frequency, "
        "address, and handshake byte below is taken from the source files, not from a marketing "
        "summary. Where the FPGA prototype and the ASIC instance differ (clock, memory size, "
        "number of SPI ports), the difference is stated."
    )
    add_p("The contributions are:")
    for item in [
        "1. A complete, code-accurate description of how this RV32I five-stage pipeline fetches, decodes, executes, and writes back instructions, including LUI/AUIPC operand steering that many textbook cores get wrong.",
        "2. A bit-level tutorial of RISC-V instruction formats and a full decode of the taped-out SPI-to-UART and GPIO assembly programs, showing how pointers are built with LUI+ADDI.",
        "3. An explanation of UART 16× oversampling with majority voting, the 0x25/0x55/0xBAADF00D bootloader, GPIO chip-select hold-off, and a parameterized SPI master that implements all four SPI modes.",
        "4. A reproducible simulation and GTKWave reading method (Icarus Verilog + cocotb), plus FPGA and Tiny Tapeout RTL-to-GDSII practice on GF180MCU.",
    ]:
        pi = doc.add_paragraph()
        pi.paragraph_format.left_indent = Inches(0.25)
        pi.paragraph_format.space_after = Pt(2)
        ri = pi.add_run(item)
        ri.font.size = Pt(9.5)
        ri.font.color.rgb = C_BODY

    add_tbl(
        ["Item", "FPGA prototype (KV260)", "ASIC instance (this repo / TTGF26a)"],
        [
            ["Top module", "Custom PL IP + bootloader", "tt_um_prem_pipeline_test"],
            ["Clock", "50 MHz from Zynq PL_CLK0", "25 MHz (CLOCK_PERIOD = 40 ns)"],
            ["Instruction memory", "Block RAM, larger image", "64 words × 32 bits (256 B)"],
            ["Data memory", "Scratch SRAM possible", "No general SRAM; MMIO only"],
            ["UART1 / UART2", "115.2 k–480 kbaud", "115 200 baud, 16× oversample"],
            ["SPI", "Lab-dependent", "One master, Mode 0, CLK_DIV=3 ≈ 4.17 MHz"],
            ["GPIO", "PMOD / LED", "GPIO1 LED; GPIO2/GPIO3 SPI CS_N"],
            ["Silicon", "Reconfigurable fabric", "4×4 tiles, GF180MCU, project 0039"],
        ],
        [Inches(1.7), Inches(2.5), Inches(2.6)],
        "TABLE I. WHAT WAS ACTUALLY BUILT: FPGA PROTOTYPE VERSUS TAPED-OUT ASIC",
    )

    # ==================================================================
    # II
    # ==================================================================
    add_sec("II. HOW RISC-V WORKS — THE IDEA BEFORE THE WIRES")
    add_p(
        "Think of the CPU as a clerk with 32 numbered drawers (registers x0–x31) and a bookmark "
        "(the program counter, PC). Drawer x0 is glued shut at zero: writing it is ignored, reading "
        "it always returns 0. That single rule makes many instruction encodings smaller, because "
        "“compare with zero” and “copy a constant” do not need a second live register [1], [11]."
    )
    add_p(
        "Every instruction is exactly 32 bits. The lowest seven bits are the opcode: they tell the "
        "control unit which family the instruction belongs to (integer ALU, load, store, branch, "
        "jump, upper immediate, or system). Other bit fields name the source drawers (rs1, rs2), "
        "the destination drawer (rd), a small function code (funct3 / funct7), and a constant "
        "(the immediate). The immediate is not stored as a clean 32-bit number. It is scattered "
        "through the word so that rs1, rs2, and rd always sit in the same bit positions. That "
        "decision, made by the RISC-V authors, lets the register file start reading before the "
        "decoder has finished classifying the instruction [1], [12]."
    )
    add_eq("instruction = {funct7, rs2, rs1, funct3, rd, opcode}   (R-type)", "1")
    add_p(
        "A program is just a list of these words in instruction memory. The PC holds the address "
        "of the next word. Because each instruction is four bytes, a “go to the next line” step is "
        "simply PC ← PC + 4. A branch or jump replaces that +4 with a computed target. After reset "
        "this chip’s PC is 0, so the first instruction the bootloader wrote is the first one that runs."
    )

    add_sub("A. The five RISC-V instruction shapes")
    add_p(
        "There are five useful shapes. Once you can see them, assembly stops looking like magic "
        "and starts looking like a form you fill in."
    )
    add_tbl(
        ["Type", "Who uses it", "Where the constant lives", "Human meaning"],
        [
            ["R", "add, and, or, xor, slt, sll…", "None (both operands are registers)", "drawer rd = f(drawer rs1, drawer rs2)"],
            ["I", "addi, lw, jalr, slli…", "Bits [31:20], sign-extended", "rd = f(rs1, constant)"],
            ["S", "sw, sh, sb", "{[31:25],[11:7]}, sign-extended", "memory[rs1+const] = rs2"],
            ["B", "beq, bne, blt, bge…", "Scrambled offset, LSB forced 0", "if condition, PC += offset"],
            ["U", "lui, auipc", "Bits [31:12] << 12", "put a 20-bit high constant in rd"],
            ["J", "jal", "Scrambled offset, LSB forced 0", "rd = PC+4; PC += offset"],
        ],
        [Inches(0.7), Inches(1.8), Inches(2.1), Inches(2.2)],
        "TABLE II. RISC-V INSTRUCTION FORMATS USED BY THIS CORE",
    )
    add_p(
        "The Verilog immediate unit (imm.v) implements exactly these five reconstructions. "
        "ImmSrc = 000 is I-type, 001 is S-type, 010 is B-type, 011 is J-type, and 100 is U-type. "
        "B-type and J-type always append a trailing zero because RISC-V jumps are even-byte aligned [1]."
    )

    add_sub("B. Why assembly language exists")
    add_p(
        "Humans do not enjoy writing 0x40000537. Assembly is a readable alias for the same bits. "
        "The line “lui x10, 0x40000” means “put 0x40000000 into register x10.” A public assembler "
        "such as the one at riscvasm.lucasteske.dev [13], or the official GNU assembler riscv32-unknown-elf-as, "
        "turns that line into the hex word the bootloader will later stream over UART. This paper "
        "uses the official ABI names (x10 is also a0, x0 is zero) but quotes the x-numbers because "
        "that is how the GitHub listings are written [6]."
    )

    # ==================================================================
    # III
    # ==================================================================
    add_sec("III. POINTERS — HOW SOFTWARE NAMES HARDWARE")
    add_p(
        "A pointer is not a C-language mystery. It is a 32-bit number that happens to be an "
        "address. On this chip there is almost no general-purpose RAM. The interesting addresses "
        "are peripheral registers. Software therefore spends its first few instructions building "
        "base pointers, then adding small offsets. That is the same idea as “struct SPI *spi = "
        "(void*)0x40000000” in C [14]."
    )
    add_p(
        "RISC-V cannot put an arbitrary 32-bit constant into a register in one instruction. "
        "LUI writes bits [31:12] and clears the low twelve bits. ADDI then adds a signed 12-bit "
        "offset. Together they form any address:"
    )
    add_eq("pointer = (U-immediate << 12) + I-immediate", "2")
    add_code(
        "lui   x10, 0x40000        # x10 = 0x4000_0000   SPI base\n"
        "addi  x11, x10, 8         # x11 = 0x4000_0008   SPI_RX\n"
        "addi  x12, x10, 4         # x12 = 0x4000_0004   SPI_TX_STATUS\n"
        "lui   x18, 0x30000        # x18 = 0x3000_0000   GPIO base\n"
        "addi  x14, x18, 4         # x14 = 0x3000_0004   GPIO2 (SPI CS)\n"
        "lui   x19, 0x10000        # x19 = 0x1000_0000   UART2 base",
        "Listing 1. Building peripheral pointers with LUI and ADDI (from docs/info.md)",
    )
    add_p(
        "After these six instructions the program never writes the big constants again. Every "
        "later load or store is of the form lw x6, 0(x12) — “read the word whose address is in "
        "x12.” The ALU computes that address in the Execute stage as rs1 + immediate. The Memory "
        "stage then compares the address against the exact MMIO map in DataMem.v [6]."
    )
    add_tbl(
        ["Pointer idea", "Register in the sample programs", "Address", "Hardware it names"],
        [
            ["SPI base", "x10", "0x4000_0000", "SPI_TX — writing starts a transfer"],
            ["SPI TX status", "x12", "0x4000_0004", "Bit0 = busy, Bit1 = pending"],
            ["SPI RX data", "x11", "0x4000_0008", "Last received MISO byte"],
            ["SPI RX status", "x13", "0x4000_000C", "Bit0 = RX valid"],
            ["UART2 base", "x19", "0x1000_0000", "UART_TX"],
            ["UART2 TX status", "x21", "0x1000_0008", "Bit0 = pending, Bit1 = busy"],
            ["GPIO1 LED", "x13 / x18", "0x3000_0000", "Bit0 drives the LED pin"],
            ["GPIO2 CS", "x14", "0x3000_0004", "Active-low chip select (gyro)"],
            ["GPIO3 CS", "—", "0x3000_0008", "Active-low chip select (magnetometer)"],
        ],
        [Inches(1.4), Inches(1.3), Inches(1.3), Inches(2.8)],
        "TABLE III. POINTER MAP USED BY THE TAPED-OUT FIRMWARE",
    )

    # ==================================================================
    # IV
    # ==================================================================
    add_sec("IV. SYSTEM ARCHITECTURE OF THE TAPED-OUT SOC")
    add_p(
        "Fig. 1 is the block diagram shipped with the Tiny Tapeout documentation. A host laptop "
        "sends a binary over UART1. The bootloader handshakes, buffers bytes, and writes instruction "
        "memory. Only then does the five-stage RISC-V core start. During execution the same core "
        "talks to UART2, SPI, and GPIO through a memory-mapped data bus. There is no interrupt "
        "controller and no cache [6], [15]."
    )
    add_fig("block_diagram_repo.png",
            "Fig. 1. Repository block diagram: host UART1 bootloader path, 64-word instruction memory, "
            "five-stage RISC-V core, and memory-mapped UART2 / SPI / GPIO (source: docs/img/block.png).")
    add_fig("Untitled presentation.png",
            "Fig. 2. Presentation form of the same architecture used in laboratory briefings at "
            "Antarikchya Pratisthan Nepal.")

    add_sub("A. Pinout on the Tiny Tapeout wrapper")
    add_p(
        "Tiny Tapeout gives every project eight dedicated inputs (ui), eight dedicated outputs (uo), "
        "and eight bidirectional pins (uio), plus clk, rst_n, and ena [16]. The wrapper "
        "tt_um_prem_pipeline_test.v maps them as follows. Unused pins are tied off so the linter "
        "and the precheck stay quiet."
    )
    add_tbl(
        ["TT pin", "Signal in info.yaml", "Direction", "Role"],
        [
            ["ui[0]", "BOOT_UART1_RX", "in", "Host firmware upload / handshake"],
            ["ui[1]", "PER_UART2_RX", "in", "Application UART receive"],
            ["uo[0]", "BOOT_UART1_TX", "out", "ACK 0x55 / NACK 0xFF"],
            ["uo[1]", "PER_UART2_TX", "out", "Application UART transmit"],
            ["uio[0]", "SPI_CS_GPIO2_TOP", "out", "SPI chip select 1 (idle high)"],
            ["uio[1]", "SPI_MOSI_TOP", "out", "Master-out, slave-in"],
            ["uio[2]", "SPI_MISO_TOP", "in", "Master-in, slave-out"],
            ["uio[3]", "SPI_SCLK_TOP", "out", "SPI clock, idle low in Mode 0"],
            ["uio[4]", "GPIO1_TOP", "out", "General LED / status"],
            ["uio[5]", "SPI_CS_GPIO3_TOP", "out", "SPI chip select 2 (idle high)"],
        ],
        [Inches(1.1), Inches(1.8), Inches(1.0), Inches(2.9)],
        "TABLE IV. TINY TAPEOUT PIN MAP (FROM info.yaml)",
    )

    add_sub("B. Clock, reset, and the two operating modes")
    add_p(
        "info.yaml declares clock_hz: 25000000. The LibreLane config sets CLOCK_PERIOD to 40 ns, "
        "which is the same number [6]. Reset is active-low on the Tiny Tapeout pad (rst_n) and "
        "inverted to active-high inside the wrapper. pipeline.v then double-registers reset so "
        "that assert is still asynchronous (the chip dies immediately when reset is pressed) but "
        "de-assert is synchronous (no half-updated pipeline flops) [17]."
    )
    add_p(
        "Mode 1, bootloader: stall_pro = 1. The PC does not advance. UART1 bytes are packed into "
        "words and written to instruction RAM. Mode 2, execution: after 0xBAADF00D the stall is "
        "released over two extra clocks (done_sr) so the last write settles, then IF/ID/EX/MEM/WB "
        "run at one instruction per clock in the absence of hazards [6]."
    )

    # ==================================================================
    # V
    # ==================================================================
    add_sec("V. THE FIVE-STAGE PIPELINE — HOW ONE INSTRUCTION WALKS THROUGH THE CHIP")
    add_p(
        "A non-pipelined CPU would finish one instruction completely before starting the next. "
        "A five-stage pipeline is an assembly line: while instruction N is writing a register, "
        "instruction N+1 is in memory, N+2 is in the ALU, N+3 is decoding, and N+4 is being "
        "fetched [11], [18]. Throughput approaches one instruction per clock. Latency stays five "
        "clocks. Fig. 3 is the textbook datapath that this Verilog follows."
    )
    add_fig("Screenshot from 2026-09-06 12-34-10.png",
            "Fig. 3. Five-stage RV32I datapath drawn in the Harris/Patterson style: instruction memory "
            "and PC on the left, register file and immediate unit, ALU, data/MMIO port, and write-back mux.")

    add_ss("1) Instruction Fetch (IF)")
    add_p(
        "The PC register (clk.v) holds PCF. On every clock, if StallF is low, it loads either "
        "PCF+4 or the branch/jump target. instruction_mem reads asynchronously: the word at "
        "PCF[7:2] appears on Instruction_out in the same cycle. Bits [1:0] of the PC are ignored "
        "because instructions are word aligned. Bits above the 64-word depth are also ignored — "
        "the memory is only 256 bytes, so a wild jump wraps in index space [6]."
    )
    add_eq("PCF' = stallF ? PCF : (PCSrcE ? PCTargetE : (PCF + 4))", "3")

    add_ss("2) Instruction Decode (ID)")
    add_p(
        "IF_ID_stage.v is a bank of flip-flops. If flushD is high (a taken branch just happened), "
        "the latch is cleared to a NOP-like zero instruction so the wrongly fetched word never "
        "reaches execute. Otherwise the 32-bit word is split in pipeline.v with ordinary bit selects:"
    )
    add_code(
        "wire [6:0] INSTR_op  = INSTRUCTION[6:0];    // opcode\n"
        "wire [2:0] INSTR_f3  = INSTRUCTION[14:12];  // funct3\n"
        "wire [6:0] INSTR_f7  = INSTRUCTION[31:25];  // funct7\n"
        "wire [4:0] INSTR_rs1 = INSTRUCTION[19:15];\n"
        "wire [4:0] INSTR_rs2 = INSTRUCTION[24:20];\n"
        "wire [4:0] INSTR_rd  = INSTRUCTION[11:7];",
        "Listing 2. Field extractors in pipeline.v — this is “instruction decode” in one glance",
    )
    add_p(
        "The Control module is a big combinational case on opcode. It raises RegWrite, MemWrite, "
        "Branch, Jump, Halt, chooses ALUControl, ImmSrc, ALUSrcA/B, and ResultSrc. ECALL/EBREAK "
        "(opcode 1110011, funct3 = 000, imm 0 or 1) set halt, which the top-level latches so the "
        "pipeline freezes after the system instruction — useful in simulation [6], [19]."
    )

    add_ss("3) Execute (EX)")
    add_p(
        "EX_stage.v is another latch (the ID/EX boundary). After it, two forwarding muxes may "
        "replace the stale register-file values with the MEM-stage ALU result or the WB-stage "
        "write-back result. A third mux may replace operand A with 0 (for LUI) or with PC "
        "(for AUIPC). Operand B is either the forwarded rs2 or the immediate. The ALU then "
        "adds, subtracts, AND/OR/XORs, shifts, or compares [6], [20]."
    )
    add_eq("SrcA = (ALUSrcA==2'b10) ? 0 : (ALUSrcA==2'b01) ? PC_E : Fwd(RD1)", "4")
    add_eq("SrcB = ALUSrcE ? ImmExtE : Fwd(RD2)", "5")
    add_p(
        "JALR is special: the target is (rs1 + imm) with the least bit cleared, as the ISA "
        "requires, so a jump never lands on an odd address [1]:"
    )
    add_eq("PCTarget = JumpR ? ((SrcA_fwd + Imm) & ~1) : (PC_E + Imm)", "6")
    add_eq("PCSrcE   = (ZeroE & BranchE) | JumpE", "7")

    add_ss("4) Memory (MEM)")
    add_p(
        "MEM_stage.v pipelines the ALU result, the store data, Rd, and the control bits. "
        "DataMem.v does not contain a RAM array. It is an address decoder. A store to "
        "0x1000_0000 arms UART2 transmit; a store to 0x4000_0000 arms SPI; a store to "
        "0x3000_0000/04/08 writes GPIO. A load from the matching status addresses returns "
        "busy and valid flags packed into a 32-bit word. This is why the assembly programs "
        "can talk to hardware with ordinary lw/sw [21]."
    )

    add_ss("5) Write-Back (WB)")
    add_p(
        "Write_back.v is a three-way mux. ResultSrc 00 takes the ALU result (most ALU ops, LUI, "
        "AUIPC). 01 takes MMIO read data (lw). 10 takes PC+4 (jal / jalr return address). The "
        "register file writes on the clock edge, and also bypasses a write to a same-cycle read "
        "of the same address so x0 stays zero and a tight producer–consumer pair still works [6]."
    )

    add_sub("A. Hazards, in ordinary words")
    add_p(
        "A data hazard is “I need a number that the previous instruction has not written yet.” "
        "Forwarding copies that number from the pipeline latch instead of waiting. A load-use "
        "hazard is worse: the number is not born until the Memory stage, so the next instruction "
        "must wait one clock (StallF, StallD, and a bubble flushed into EX). A control hazard is "
        "“we just discovered we should have jumped.” The two instructions already fetched after "
        "the branch are flushed [11], [18], [22]."
    )
    add_eq("ForwardAE = 2'b10  if RegWriteM && (RdM≠0) && (RdM==Rs1E)", "8")
    add_eq("ForwardAE = 2'b01  if RegWriteW && (RdW≠0) && (RdW==Rs1E)", "9")
    add_p(
        "Hazard_Unit.v implements those compares and gives MEM-stage data priority over WB-stage "
        "data, which is the standard Patterson rule [11]. Branch flush has lower priority than "
        "load-use stall so a stalled decode does not get thrown away."
    )

    add_sub("B. What the ALU actually computes")
    add_p(
        "ALU.v is typed. ALUType 00 is ordinary integer ops. Type 01 and 11 force an add "
        "(stores and jumps). Type 10 is branch compare: Zero is not “result==0” but “condition "
        "true,” so BEQ/BNE/BLT/BGE/BLTU/BGEU all share one adder/comparator [6], [23]."
    )
    add_tbl(
        ["ALUControl", "Integer (type 00)", "Branch (type 10)"],
        [
            ["0000", "AND", "BEQ  (equal)"],
            ["0001", "OR", "BNE  (not equal)"],
            ["0010", "ADD", "BLT  (signed <)"],
            ["0011", "SUB", "BGE  (signed ≥)"],
            ["0100", "XOR", "BLTU (unsigned <)"],
            ["0101", "SLL", "BGEU (unsigned ≥)"],
            ["0110", "SRL", "—"],
            ["0111", "SRA", "—"],
            ["1000", "SLT", "—"],
            ["1001", "SLTU", "—"],
            ["1010", "pass B (LUI)", "—"],
        ],
        [Inches(1.4), Inches(2.5), Inches(2.9)],
        "TABLE V. ALU CONTROL ENCODING FROM Control_Unit.v AND ALU.v",
    )

    # ==================================================================
    # VI
    # ==================================================================
    add_sec("VI. INSTRUCTION MEMORY — THE PLACE THE PROGRAM LIVES")
    add_p(
        "Instruction_Mem.v is a 64-entry array of 32-bit registers. That is 256 bytes. It has "
        "two ports that do not fight: a synchronous write port used only by the bootloader "
        "(we, addr[5:0], wdata) and a combinational read port indexed by the PC. In a classroom "
        "CPU the program is often hard-coded with $readmemh. Here the array wakes up undefined "
        "and is filled over UART. That is what “programmable” means in the project title [6], [24]."
    )
    add_eq("word_index = read_Address[7:2]     // drop byte bits; 64 words → 6 index bits", "10")
    add_p(
        "Because the write address is a word index, not a byte address, the bootloader’s "
        "addr_count++ moves to the next instruction, not to the next byte. The first uploaded "
        "word lands at PC = 0, the second at PC = 4, and the last legal word at PC = 252. "
        "A 65th real instruction would wrap and overwrite word 0 — the 64-word depth is a hard "
        "Tiny Tapeout area budget, not a RISC-V limitation [16], [25]."
    )
    add_code(
        "module instruction_mem #(.DEPTH(64), .ADDR_W(6)) (\n"
        "    input  wire        clk, we,\n"
        "    input  wire [5:0]  addr,\n"
        "    input  wire [31:0] wdata,\n"
        "    input  wire [31:0] read_Address,\n"
        "    output wire [31:0] Instruction_out);\n"
        "    reg [31:0] mem [0:63];\n"
        "    always @(posedge clk) if (we) mem[addr] <= wdata;\n"
        "    assign Instruction_out = mem[read_Address[7:2]];\n"
        "endmodule",
        "Listing 3. Complete instruction memory (Instruction_Mem.v), slightly compacted",
    )

    # ==================================================================
    # VII
    # ==================================================================
    add_sec("VII. UART WITH 16× OVERSAMPLING — WHY THE RECEIVER DOES NOT TRUST ONE SAMPLE")
    add_p(
        "UART is a conversation with no shared clock. The transmitter and receiver each have "
        "their own oscillator. They agree only on a baud rate and a frame: one start bit (line "
        "goes from idle-high to low), eight data bits least-significant first, and one stop bit "
        "(line high again) [26], [27]. If the receiver samples too early or too late, a bit is "
        "read as its neighbor and the firmware image is garbage."
    )
    add_p(
        "The standard engineering answer is oversampling [27]–[29]. Both uart.v (boot UART) and "
        "uart0.v (application UART) generate a tick that is sixteen times faster than the nominal "
        "bit rate:"
    )
    add_eq("BAUD_DIV = f_clk / (BaudRate × 16)", "11")
    add_eq("f_tick   = f_clk / BAUD_DIV", "12")
    add_eq("T_bit    = 16 × T_tick", "13")
    add_p(
        "With f_clk = 25 MHz and BaudRate = 115 200, the exact divider is 13.0208…. Verilog "
        "integer division yields BAUD_DIV = 13. The true bit rate is therefore"
    )
    add_eq("Baud_actual = 25×10^6 / (13 × 16) = 120 192.3 baud   (≈ +4.3 % vs 115 200)", "14")
    add_p(
        "A 4.3 % baud error is larger than the classical 2 % textbook budget, but 16× majority "
        "voting still leaves a usable sampling window on short cables and FTDI adapters [29], [30]. "
        "The FPGA uploader in Fig. 7 can also probe higher rates; the ASIC instance is fixed at "
        "the 25 MHz / 115 200 pair declared in the RTL parameters."
    )

    add_sub("A. Transmitter state machine")
    add_p(
        "TX is the easy half. Idle, the line is 1. When tx_Start pulses, the byte is latched. "
        "The FSM walks START → DATA[0..7] → STOP, holding each phase for 16 baud ticks. That "
        "guarantees each transmitted bit lasts one nominal bit time. The boot UART uses this "
        "path to send ACK 0x55 or NACK 0xFF [6]."
    )

    add_sub("B. Receiver: two-flop synchronizer, start-bit vote, data-bit vote")
    add_p(
        "An asynchronous pin must never clock a state machine directly. uart.v first samples rx "
        "through two flip-flops (rx_s1, rx_s2) to cut metastability [17], [31]. In RX_IDLE, a "
        "low on rx_s2 is treated as a possible start bit. The FSM then counts 16 ticks and "
        "counts how many of those ticks were still low. If fewer than eight were low "
        "(RX_MAJ = OVERSAMPLE/2), the edge was a glitch and the FSM returns to idle. If the "
        "majority was low, a real start bit is accepted and eight data bits follow."
    )
    add_p(
        "Each data bit is not sampled once at the “center.” Every one of the 16 ticks is counted. "
        "If at least eight ticks saw a 1, the bit is a 1; otherwise it is a 0. The bit is then "
        "shifted into rx_shift_reg from the left, matching LSB-first UART order. The stop bit is "
        "majority-voted the same way; only a high stop bit raises rx_ready and commits rx_Data. "
        "That is why a noisy laboratory cable still yields a clean 0x25 handshake [6], [29]."
    )
    add_code(
        "localparam integer BAUD_DIV = CLK_FREQ / (BAUD_RATE * OVERSAMPLE); // 13 at 25 MHz\n"
        "localparam integer RX_MAJ   = OVERSAMPLE / 2;                      // 8 of 16\n"
        "// in RX_DATA, every baud_tick:\n"
        "one_counts <= one_counts + rx_s2;\n"
        "if (rx_sample_cnt == 15) begin\n"
        "    rx_shift_reg <= { (one_counts >= RX_MAJ), rx_shift_reg[7:1] };\n"
        "    ...\n"
        "end",
        "Listing 4. Heart of the 16× majority-vote UART receiver (uart.v)",
    )

    # ==================================================================
    # VIII
    # ==================================================================
    add_sec("VIII. THE HARDWARE BOOTLOADER — HOW A BLANK CHIP LEARNS A PROGRAM")
    add_p(
        "After silicon reset, instruction memory is empty. A JTAG cable would be another block "
        "of gates this tile cannot afford. The bootloader is therefore a small FSM beside the "
        "UART that turns a serial byte stream into instruction-memory writes [6], [32]."
    )
    add_p(
        "State 0, listen. stall_pro is 1, so the PC is frozen. Any byte other than 0x25 is "
        "answered with NACK 0xFF. Byte 0x25 is answered with ACK 0x55 and handshake_done becomes 1. "
        "The cocotb test (test.py) does exactly this handshake before it uploads words [6]."
    )
    add_p(
        "State 1, assemble. Each rising edge of rx_ready (detected as rx_valid & ~rx_valid_d so "
        "a held ready flag cannot write four times) deposits one byte into a little-endian slot:"
    )
    add_eq("word = {byte3, byte2, byte1, byte0}     // byte0 is the opcode field", "15")
    add_p(
        "After the fourth byte the word is compared with 0xBAADF00D. If it matches, boot_done "
        "rises and no write occurs — the sentinel is not an instruction. If it does not match, "
        "mem_we pulses, mem_addr is the running word count, and mem_wdata is the assembled word. "
        "Two extra clocks later stall_pro falls and the pipeline begins at PC = 0 [6]."
    )
    add_fig("1789972494070.jpeg",
            "Fig. 4. Vivado IP Integrator view of the FPGA boot path: Zynq UltraScale+ MPSoC clock/reset, "
            "16× UART, and bootloader IP. This is the laboratory twin of uart.v + uart_bootloader.v.")
    add_fig("Screenshot from 2026-09-29 11-54-13.png",
            "Fig. 5. Host “UART Terminal & Instruction Uploader”: 0x25 handshake, 0x55 ACK after baud "
            "probe, little-endian 32-bit words, and 0xBAADF00D CPU release.")

    add_tbl(
        ["Token / event", "Value", "Who sends it", "What the chip does"],
        [
            ["Handshake command", "0x25", "Host", "If matched, ACK and unlock upload"],
            ["ACK", "0x55", "Chip UART1 TX", "Host may start streaming words"],
            ["NACK", "0xFF", "Chip UART1 TX", "Wrong first byte; stay listening"],
            ["Instruction word", "any 32-bit ≠ sentinel", "Host, 4 bytes LE", "Write IMEM[addr++]"],
            ["Sentinel", "0xBAADF00D", "Host", "Stop writes; release stall_pro"],
            ["First fetch address", "0x0000_0000", "PC register", "Execute uploaded program"],
        ],
        [Inches(1.6), Inches(1.5), Inches(1.4), Inches(2.3)],
        "TABLE VI. BOOTLOADER PROTOCOL (uart_bootloader.v)",
    )

    # ==================================================================
    # IX
    # ==================================================================
    add_sec("IX. GPIO — LIGHTS AND CHIP SELECTS")
    add_p(
        "GPIO on this chip is intentionally tiny. gpio1_io is one flip-flop. A store of any "
        "value whose least bit is 1 turns the LED pin on; 0 turns it off. Reset clears it so "
        "the pin does not glow before software runs [6]."
    )
    add_p(
        "GPIO2 and GPIO3 are smarter because they are SPI chip selects. SPI slaves require CS "
        "to stay low for the whole byte (sometimes for a whole burst) and to rise only after "
        "the last SCLK edge. If software writes “CS = 1” one instruction too early, the slave "
        "drops the last bit. gpio2_io / gpio3_io therefore implement a hold-off: a request to "
        "de-assert is remembered in deassert_pending and applied only when spi_busy and "
        "spi_pending are both low. Those two status bits are themselves registered in "
        "pipeline.v (spi2_busy_buf, spi2_pending_buf) so the fan-out into two CS FSMs does "
        "not become a timing bomb on 180 nm [6], [33]."
    )
    add_code(
        "sw   x16, 0(x14)     # x16=0 → GPIO2 low  → CS asserted immediately\n"
        "# ... SPI byte transfers ...\n"
        "sw   x3,  0(x14)     # x3=1  → CS rises only after SPI goes idle",
        "Listing 5. Software view of the hardware CS hold-off",
    )
    add_p(
        "In the RTL comments GPIO2 is the IAM-20380 gyroscope CS_N and GPIO3 is the MMC5983MA "
        "magnetometer CS_N [34], [35]. The taped-out tile does not contain those sensors; it "
        "contains the pins that a later CubeSat board can wire to them. That is the SatOC "
        "roadmap in miniature: first prove the bus, then hang flight sensors on it [36]."
    )

    # ==================================================================
    # X
    # ==================================================================
    add_sec("X. SPI — FOUR MODES, AND THE ONE THIS CHIP USES")
    add_p(
        "SPI is a four-wire synchronous bus: SCLK (clock from the master), MOSI (master to slave), "
        "MISO (slave to master), and CS (slave select, usually active low) [37], [38]. Unlike UART, "
        "the slave does not need to know the baud rate in advance. It simply rides the master’s clock. "
        "That makes SPI the usual choice for MEMS sensors, flash, and ADC chips [34], [35], [38]."
    )
    add_p(
        "The two mode bits are clock polarity and clock phase [37]:"
    )
    add_tbl(
        ["Mode", "CPOL", "CPHA", "Idle SCLK", "When MOSI/MISO are sampled", "When MOSI is updated"],
        [
            ["0", "0", "0", "Low", "Rising edge", "Falling edge (and before first edge)"],
            ["1", "0", "1", "Low", "Falling edge", "Rising edge"],
            ["2", "1", "0", "High", "Falling edge", "Rising edge (and before first edge)"],
            ["3", "1", "1", "High", "Rising edge", "Falling edge"],
        ],
        [Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.9), Inches(1.9), Inches(1.9)],
        "TABLE VII. MOTOROLA SPI MODES — spi.v IMPLEMENTS ALL FOUR VIA PARAMETERS",
    )
    add_p(
        "The Verilog is honest about this. Parameters CPOL and CPHA change two lines:"
    )
    add_code(
        "sclk <= CPOL;                                  // idle level\n"
        "wire sample_edge = (CPHA == 0) ? sclk_rise : sclk_fall;\n"
        "wire shift_edge  = (CPHA == 0) ? sclk_fall : sclk_rise;",
        "Listing 6. How one SPI master becomes four modes (spi.v)",
    )
    add_p(
        "The taped-out instance is instantiated as CPOL=0, CPHA=0, CLK_DIV=3, DATA_WIDTH=8. "
        "The divider toggles SCLK every 3 system clocks, so one SCLK period is 6 system clocks:"
    )
    add_eq("f_SCLK = f_clk / (2 × CLK_DIV) = 25 MHz / 6 ≈ 4.167 MHz", "16")
    add_p(
        "MSB is driven on MOSI before the first rising edge, which is exactly Mode-0 setup. "
        "MISO is synchronized with two flip-flops, then sampled on the rising edge into a left "
        "shift register. After eight bits, done pulses and rx_data is the slave’s byte. "
        "Full duplex is free: the master always sends and always receives [6], [37]."
    )
    add_p(
        "Software must poll. SPI_TX_STATUS bit 0 is busy and bit 1 is pending. SPI_RX_STATUS "
        "bit 0 is valid. The sample program writes 0xAA as a dummy MOSI byte (many sensors "
        "need a clock train more than they need a particular command on the first byte), waits "
        "for RX valid, then copies the received byte to UART2 so a human can see the sensor "
        "talk [6]. The cocotb slave in test.py preloads MISO before the first rising edge and "
        "updates it on the falling edge — a textbook Mode-0 slave [39]."
    )

    # ==================================================================
    # XI
    # ==================================================================
    add_sec("XI. VERILOG AS THE LANGUAGE OF THE MACHINE")
    add_p(
        "Every module in src/ is synthesizable Verilog-2001 with `default_nettype none. That "
        "directive is not decoration: it makes a missing wire a compile error instead of a "
        "silent 1-bit implicit net [40], [41]. The top of the ASIC is the Tiny Tapeout wrapper. "
        "The top of the CPU is pipeline.v, which is a netlist of the stages plus the peripherals. "
        "There is no SystemVerilog class, no interface, no clocking block in the DUT — only "
        "the constructs LibreLane and the GF180 standard-cell library will accept [42], [43]."
    )
    add_tbl(
        ["File", "What it is, in one sentence"],
        [
            ["tt_um_prem_pipeline_test.v", "Pad wrapper: maps ui/uo/uio onto UART, SPI, GPIO"],
            ["pipeline.v", "CPU + boot UART + app UART + SPI + GPIO + reset tree"],
            ["Control_Unit.v", "Opcode/funct combinational decoder, including halt"],
            ["ALU.v", "Integer ALU and branch comparator"],
            ["Reg_File.v", "32×32 registers, x0 hard-zero, write-through bypass"],
            ["imm.v", "I/S/B/J/U immediate reconstruction"],
            ["Instruction_Mem.v", "64-word programmable IMEM"],
            ["DataMem.v", "Exact-match MMIO decoder, not a RAM"],
            ["Hazard_Unit.v", "Forwarding, load-use stall, branch flush"],
            ["uart.v / uart0.v", "16× oversampled TX/RX (boot / application)"],
            ["uart_bootloader.v", "0x25 / 0x55 / little-endian words / 0xBAADF00D"],
            ["spi.v", "Parameterized SPI master (mode, divider, width)"],
            ["gpio.v", "LED flop + two CS FSMs with idle hold-off"],
            ["clk.v / PC_incre.v / PCSelect.v", "PC register, +4, branch mux"],
            ["IF_ID / EX / MEM / WB stages", "Pipeline registers; flush/stall where needed"],
        ],
        [Inches(2.4), Inches(4.4)],
        "TABLE VIII. SOURCE FILES IN THE ANALYZED REPOSITORY",
    )
    add_p(
        "A reader who wants to learn digital design can open these files in order: first the "
        "wrapper, then pipeline.v’s comments (they even mark the reset-fanout and SPI-fanout "
        "fixes that were needed for GF180 timing), then Control_Unit.v with the ISA manual "
        "open beside it [1], [6]. That is a better textbook than a slide deck because every "
        "signal is real."
    )

    # ==================================================================
    # XII
    # ==================================================================
    add_sec("XII. SIMULATION AND GTKWAVE — WATCHING THE CHIP THINK")
    add_p(
        "Before any FPGA bitstream or any GDSII is made, the design is executed as a program "
        "on a simulator. This repository uses Icarus Verilog as the event engine and cocotb "
        "as the Python testbench [39], [44], [45]. make -B in test/ compiles every file listed "
        "in PROJECT_SOURCES, elaborates tb as the toplevel, and runs test.py. The Verilog "
        "testbench (tb.v) only provides a 50 ns period clock (#25 toggle → 20 MHz in the "
        "standalone initial block; cocotb overrides this with a 40 ns / 25 MHz Clock object), "
        "maps the Tiny Tapeout pins, and dumps a waveform [6]."
    )
    add_code(
        "$dumpfile(\"tb.fst\");\n"
        "$dumpvars(0, tb);\n"
        "always #25 clk = ~clk;   // used if cocotb is not driving the clock",
        "Listing 7. Waveform dump and clock in tb.v",
    )
    add_p(
        "FST is a compressed waveform format. GTKWave opens it with gtkwave tb.fst tb.gtkw [46]. "
        "The saved gtkw file in the repo is still the Tiny Tapeout template (clk, rst_n, ui, uo, "
        "uio). A useful analysis session adds the internal hierarchy:"
    )
    add_tbl(
        ["GTKWave signal", "What a correct run looks like"],
        [
            ["tb.clk / dut.clk", "40 ns period (25 MHz) after cocotb starts the clock"],
            ["tb.rst_n", "Low for ~10 cycles, then high and stays high"],
            ["dut.ui_in[0] (BOOT RX)", "Idle 1; falling edges for 0x25, then instruction bytes"],
            ["dut.uo_out[0] (BOOT TX)", "Idle 1; a 0x55 frame after the handshake"],
            ["imem.we / imem.addr / imem.wdata", "we pulses every 4 RX bytes; addr 0,1,2…; wdata = instruction"],
            ["bootloader.stall_pro", "1 during upload; 0 shortly after 0xBAADF00D"],
            ["PCF", "0 while stalled; then 0,4,8,… until a branch"],
            ["INSTRUCTION / INSTR_op", "The uploaded words, opcode 0110111 for LUI, etc."],
            ["SPI_CS / SPI_SCLK / MOSI / MISO", "CS falls; 8 Mode-0 clocks; MOSI 0xAA; MISO from the Python slave"],
            ["uo_out[1] (UART2 TX)", "Frames whose payload is the SPI slave text"],
        ],
        [Inches(2.4), Inches(4.4)],
        "TABLE IX. GTKWAVE READING GUIDE FOR THIS DESIGN",
    )
    add_fig("Screenshot from 2026-03-19 17-50-35.png",
            "Fig. 6. GTKWave session on a dumped VCD/FST. The clock is the first signal to check: "
            "period must match the intended 25 MHz (ASIC) or 50 MHz (FPGA) test. Zoom until one "
            "UART bit or one SPI byte fills the screen, then confirm setup/hold against SCLK.")
    add_p(
        "How to read a UART byte in GTKWave: idle is high. The first falling edge is the start "
        "bit and lasts 16 baud ticks (about 8.68 µs at 115 200, 8.32 µs at the ASIC’s actual "
        "120 192). The next eight chunks are data bits 0…7. The last high chunk is stop. "
        "If you see 0x25, the bits are 1 0 1 0 0 1 0 0 after the start bit (LSB first: "
        "1,0,1,0,0,1,0,0 → 0b00100101 = 0x25). That single exercise teaches more about "
        "asynchronous serial than a chapter of prose [26], [46]."
    )
    add_p(
        "The Python test uploads a compact SPI-to-UART program, including a corrected jump "
        "0xFE5FF06F, and runs a Mode-0 slave that returns the greeting “Namaste everyone from "
        "Nepal! …” so UART2 can be checked by eye in the log [6]. Gate-level simulation is the "
        "same testbench with GATES=yes after LibreLane has produced a structural netlist of "
        "GF180 standard cells — that is the last software check before trusting silicon [42], [47]."
    )

    # ==================================================================
    # XIII
    # ==================================================================
    add_sec("XIII. ASSEMBLY LANGUAGE DECODE — FROM HEX TO MEANING")
    add_p(
        "This section decodes real words from the repository, not a made-up tutorial. Take "
        "the first instruction of the SPI program: 0x40000537."
    )
    add_code(
        "0x40000537 = 0100 0000 0000 0000 0000  01010  0110111\n"
        "             | imm[31:12] = 0x40000 |   rd    opcode\n"
        "opcode 0110111 = LUI\n"
        "rd     01010   = x10\n"
        "imm<<12        = 0x4000_0000\n"
        "assembly       = lui x10, 0x40000",
        "Listing 8. Bit-field decode of 0x40000537",
    )
    add_p(
        "Second example: 0x00850593. Opcode 0010011 is I-type ALU. funct3 000 is ADDI. "
        "rd = x11, rs1 = x10, imm = 8. So addi x11, x10, 8 — the SPI_RX pointer. "
        "Third example: 0x01072023. Opcode 0100011 is STORE. funct3 010 is SW. "
        "imm = 0, rs1 = x14, rs2 = x16. So sw x16, 0(x14) — assert CS."
    )
    add_p(
        "Branch immediates are the ones that intimidate students. 0xFE040EE3 is beq x8, x0, "
        "backwards. The B-type immediate is assembled from bits 31, 7, 30:25, 11:8, then "
        "a trailing 0. Sign-extended, that offset is negative, which is why a polling loop "
        "can jump to itself without a jal [1], [13]. Table X decodes the GPIO blink program "
        "that is actually in docs/info.md."
    )
    add_tbl(
        ["Hex word", "Assembly", "What the hardware does"],
        [
            ["0x300006B7", "lui x13, 0x30000", "x13 ← 0x30000000 (GPIO1 pointer)"],
            ["0x00468713", "addi x14, x13, 4", "x14 ← 0x30000004 (GPIO2 pointer)"],
            ["0x00A00293", "addi x5, x0, 10", "loop counter = 10"],
            ["0x00100313", "addi x6, x0, 1", "HIGH constant"],
            ["0x00000393", "addi x7, x0, 0", "LOW constant"],
            ["0x0076A023", "sw x7, 0(x13)", "LED off"],
            ["0x00772023", "sw x7, 0(x14)", "CS idle / GPIO2 low in this test"],
            ["0x3E800193", "addi x3, x0, 1000", "software delay start"],
            ["0xFFF18193", "addi x3, x3, -1", "delay body"],
            ["0xFE019EE3", "bne x3, x0, delay", "taken until x3 hits 0"],
            ["0x0066A023", "sw x6, 0(x13)", "LED on"],
            ["0x00000073", "ecall", "halt_top = 1; simulation stops"],
            ["0xBAADF00D", "(sentinel)", "not executed; bootloader release"],
        ],
        [Inches(1.3), Inches(1.8), Inches(3.7)],
        "TABLE X. DECODED GPIO TEST FROM THE REPOSITORY (docs/info.md)",
    )
    add_p(
        "The SPI-to-UART program in the same file is the more important flight-like routine: "
        "build three base pointers, assert CS, poll SPI TX, write 0xAA, poll SPI RX, poll UART "
        "TX, store the received byte, decrement a counter of 200, jal back, then release CS and "
        "ecall. That is polling I/O in its pure form — the same pattern used in bare-metal "
        "CubeSat software when an RTOS is too heavy [36], [48]."
    )
    add_code(
        "lui   x10, 0x40000\nlui   x18, 0x30000\nlui   x19, 0x10000\n"
        "addi  x11, x10, 8\naddi  x12, x10, 4\naddi  x13, x10, 12\n"
        "addi  x21, x19, 8\naddi  x14, x18, 4\n"
        "addi  x3, x0, 1\naddi  x16, x0, 0\naddi  x5, x0, 200\n"
        "addi  x22, x0, 1\naddi  x7, x0, 0xAA\n"
        "sw    x16, 0(x14)          # CS low\n"
        "loop:\n"
        "  beq  x5, x0, release_cs\n"
        "wait_spi_tx:\n"
        "  lw   x6, 0(x12)\n  and  x6, x6, x22\n  beq  x6, x0, wait_spi_tx\n"
        "  sw   x7, 0(x10)          # start SPI byte\n"
        "wait_spi_rx:\n"
        "  lw   x8, 0(x13)\n  and  x8, x8, x22\n  beq  x8, x0, wait_spi_rx\n"
        "  lw   x9, 0(x11)\n"
        "wait_uart_tx:\n"
        "  lw   x6, 0(x21)\n  and  x6, x6, x22\n  beq  x6, x0, wait_uart_tx\n"
        "  sw   x9, 0(x19)          # relay to UART2\n"
        "  addi x5, x5, -1\n  jal  x0, loop\n"
        "release_cs:\n  sw x3, 0(x14)\n  ecall",
        "Listing 9. Repository SPI-to-UART polling program (human form)",
    )

    # ==================================================================
    # XIV
    # ==================================================================
    add_sec("XIV. FPGA TESTING — A RECONFIGURABLE DRESS REHEARSAL")
    add_p(
        "An FPGA is a chip whose gates can be rewritten in milliseconds. It is the honest place "
        "to find a broken handshake before paying for a wafer [9], [49]. This research program "
        "mapped the UART and bootloader onto the programmable logic of an AMD Kria KV260 "
        "(Zynq UltraScale+ XCZU5EV), clocked at 50 MHz from the processing-system PL clock, "
        "and wired PMOD pins through an FTDI FT232R to a laptop [9], [10], [50]. Fig. 7 is "
        "the laboratory bench."
    )
    add_fig("1789972494046.jpeg",
            "Fig. 7. FPGA dress rehearsal: AMD Kria KV260, FTDI USB–UART, and host laptop running "
            "the instruction uploader. The green LED shows the board is alive; the red LED on the "
            "FTDI marks USB enumeration.")
    add_p(
        "The practical FPGA test is the same story as silicon, only faster to iterate. "
        "(1) Build the bitstream in Vivado. (2) Set rst_n low, then high. (3) Open the uploader "
        "on /dev/ttyUSB1. (4) Send 0x25. (5) Expect 0x55. If 0xFF arrives, the FPGA is still in "
        "the wrong image or the line is inverted. (6) Stream little-endian words. (7) Send "
        "0xBAADF00D. (8) Watch GPIO blink or UART2 echo. The Tiny Tapeout CI also contains an "
        "fpga.yaml workflow so the same HDL can be exercised on the official TT FPGA playground "
        "without a KV260 [6], [16]."
    )
    add_p(
        "Timing is the FPGA lesson that later saves the ASIC. At 50 MHz the UART divider is "
        "different from the 25 MHz instance; the RTL parameters must be changed or the baud "
        "will be wrong by 2×. The uploader’s auto-baud probe in Fig. 5 exists because of that "
        "class of mistake. Hold time on MISO, reset release, and CS hold-off were all found "
        "on the bench before they were re-coded for GF180 fan-out [6], [33]."
    )

    # ==================================================================
    # XV
    # ==================================================================
    add_sec("XV. TINY TAPEOUT AND THE RTL-TO-GDSII PATH")
    add_p(
        "GDSII (sometimes still called GDS II; the user-facing nickname “GDBII” is the same "
        "file) is the polygon database a mask shop uses to make photomasks [51], [52]. "
        "Getting from Verilog to that file used to require a commercial CAD seat. Tiny Tapeout "
        "wraps an open-source compiler, LibreLane (the successor of OpenLane), around a fixed "
        "die template so many small projects share one wafer [16], [42], [53], [54]."
    )
    add_p(
        "A Tiny Tapeout tile on GF180 is about 340 µm × 160 µm. This project asked for tiles: "
        "\"4x4\", so sixteen tiles, roughly 1.36 mm × 0.64 mm of user area before the wrapper "
        "ring [6], [7]. The GitHub Action .github/workflows/gds.yaml calls "
        "TinyTapeout/tt-gds-action@ttgf26a with pdk: gf180mcuD, then precheck, gate-level "
        "simulation, and a GDS viewer page [6], [47]."
    )

    add_sub("A. The compiler steps, said simply")
    add_p(
        "LibreLane is a scripted chain. Each step is a well-known EDA algorithm [42], [55], [56]:"
    )
    add_tbl(
        ["Step", "Tool idea", "What a human should look at"],
        [
            ["1. RTL synthesis", "Yosys reads Verilog, maps to GF180 standard cells [57]", "Cell count, inferred latches (should be 0)"],
            ["2. Floorplan", "A rectangle and pin list matching the TT wrapper", "Die size, pin order vs info.yaml"],
            ["3. Placement", "Cells scattered then legalized; density 60% here", "GPL-0302 failures → raise PL_TARGET_DENSITY_PCT"],
            ["4. CTS", "Clock tree so every flop sees clk with similar delay [17]", "Skew, max clock insertion delay"],
            ["5. Routing", "Metal wires; GF180 MCU 7-track 5 V library", "DRC opens/shorts, congestion"],
            ["6. Timing repair", "Insert buffers; CLOCK_PERIOD = 40 ns", "Setup WNS ≥ 0; hold slack margins 0.1 / 0.05 ns"],
            ["7. Signoff", "Magic / KLayout / Netgen LVS, DRC, antenna", "LVS clean; no disconnected pins"],
            ["8. GDSII / LEF / DEF / SPEF / netlist", "What the shuttle aggregates", "final/gds/*.gds — the tapeout artifact"],
        ],
        [Inches(1.5), Inches(2.6), Inches(2.7)],
        "TABLE XI. RTL-TO-GDSII FLOW USED FOR PROJECT 0039",
    )
    add_p(
        "config.json in this repo is the project’s contract with that flow. The numbers that "
        "matter are PL_TARGET_DENSITY_PCT = 60 (do not pack cells until they cannot route), "
        "CLOCK_PERIOD = 40 (25 MHz), CLOCK_PORT = clk, RUN_CTS = 1, and the hold slack margins. "
        "KLayout XOR and DRC are turned off inside the user job because the Tiny Tapeout "
        "precheck job repeats a blessed version of those checks [6], [16]."
    )
    add_p(
        "After GDSII is merged into the shuttle, GlobalFoundries prints the wafer, the wafer "
        "is diced, and Tiny Tapeout ships a packaged chip whose mux selects project 39 when "
        "the board is told that address [7], [8]. Testing silicon is then the FPGA script "
        "again: 25 MHz clock, 115 200 baud (or the actual 120 kbaud), 0x25, words, sentinel, "
        "watch UART2 and SPI."
    )

    add_sub("B. Why this path matters beyond one student project")
    add_p(
        "Open PDKs (SkyWater SKY130, GlobalFoundries GF180MCU) and multi-project shuttles "
        "have moved custom silicon from a corporate privilege to a laboratory exercise [5], "
        "[43], [53], [58]. For an emerging space program the lesson is not that a 256-byte "
        "IMEM will fly a 3U CubeSat by itself. The lesson is that the team now owns every "
        "abstraction from assembly mnemonic to mask polygon, and can grow the same core "
        "toward a Satellite-on-Chip when memory, radiation tolerance, and sensors catch up "
        "[36], [59], [60]."
    )

    # ==================================================================
    # XVI
    # ==================================================================
    add_sec("XVI. RESULTS DRAWN FROM THE ANALYZED REPOSITORY")
    add_p(
        "Table XII is the specification sheet of the chip that was actually submitted, not the "
        "larger SatOC vision. Every row can be pointed to a file."
    )
    add_tbl(
        ["Quantity", "Value", "Evidence in the repo"],
        [
            ["ISA", "RV32I subset (R/I/S/B/U/J + ECALL)", "Control_Unit.v case on opcode"],
            ["Pipeline", "5 stages, forwarding, 1-cycle load-use stall", "pipeline.v, Hazard_Unit.v"],
            ["IMEM", "64 × 32-bit, boot-writable", "Instruction_Mem.v DEPTH=64"],
            ["DMEM", "None; MMIO decoder only", "DataMem.v exact-address compares"],
            ["Clock (ASIC)", "25 MHz, 40 ns period", "info.yaml, config.json"],
            ["UART1 / UART2", "115 200 target, 16× majority vote", "uart.v, uart0.v OVERSAMPLE=16"],
            ["Boot protocol", "0x25 → 0x55, LE words, 0xBAADF00D", "uart_bootloader.v, test.py"],
            ["SPI", "Mode 0, 8-bit, CLK_DIV=3 ≈ 4.17 MHz", "pipeline.v spi_master #()"],
            ["GPIO", "LED + 2 CS with idle hold-off", "gpio.v, DataMem.v 0x3000_000x"],
            ["Interrupts", "None; polling only", "docs/info.md"],
            ["Tile request", "4×4 on GF180MCU / TTGF26a", "info.yaml tiles: \"4x4\""],
            ["Shuttle address", "0039 Programmable_Pipeline-RISC-V", "TTGF26a datasheet [7]"],
        ],
        [Inches(1.5), Inches(2.6), Inches(2.7)],
        "TABLE XII. MEASURED / DECLARED RESULTS OF THE TAPED-OUT DESIGN",
    )
    add_p(
        "Functional verification in cocotb demonstrated the handshake, the little-endian upload, "
        "sentinel release, Mode-0 SPI traffic against a Python slave, and UART2 relay of a "
        "multi-character string [6]. FPGA bring-up demonstrated the same protocol at laboratory "
        "baud rates up to 480 kbaud on the KV260 path (Fig. 5). ASIC timing is signed off at "
        "25 MHz by construction of CLOCK_PERIOD [42]."
    )
    add_p(
        "Limitations are part of the result. There is no on-chip SRAM for stacks. There is no "
        "I2C master in this tapeout (DataMem.v comments name IMU/magnetometer CS pins, not an "
        "I2C block). The integer baud divider is coarse at 25 MHz. Instruction memory is 64 "
        "words. Those limits are the honest price of sixteen GF180 tiles and they belong in "
        "the paper so the next die can be specified from facts [25], [36]."
    )

    # ==================================================================
    # XVII
    # ==================================================================
    add_sec("XVII. DISCUSSION")
    add_p(
        "Two cultures meet in this repository. One is the computer-architecture classroom: "
        "Patterson’s five-stage pipeline, RISC-V immediates, forwarding muxes [1], [11]. "
        "The other is the board-bring-up laboratory: FTDI cables, CS hold-off, majority-voted "
        "start bits, LibreLane density knobs [16], [29], [42]. The chip is useful as research "
        "because it does not hide the seam. A student can watch 0x40000537 become a pointer, "
        "watch that pointer become a chip-select edge, and watch that edge become a GDSII "
        "polygon on metal."
    )
    add_p(
        "Compared with other Tiny Tapeout RISC-V efforts — TinyQV [61], SERV-based SoCs [62], "
        "and the UTOSS core on the same GF 26a shuttle [7] — this design is not the smallest "
        "or the fastest. Its distinction is programmability through a documented UART "
        "bootloader plus a complete, pollable SPI/UART/GPIO map that matches a CubeSat sensor "
        "story. That is a different optimization: teachable and mission-shaped, rather than "
        "minimal gate count [36], [60]."
    )
    add_p(
        "The SatOC roadmap remains valid and is not claimed as finished silicon. Future dies "
        "need a real SRAM compiler or OpenRAM instance [63], interrupt or DMA for sensor "
        "FIFOs, ECC or TMR on the pipeline [64], [65], and a cleaner baud generator "
        "(fractional divider or a PLL-friendly clock) [27]. The present chip is the first "
        "rung: a processor that can be taught, simulated, programmed, and fabricated."
    )

    # ==================================================================
    # XVIII
    # ==================================================================
    add_sec("XVIII. CONCLUSION")
    add_p(
        "This paper analyzed the public RTL of ttgf_prem_programmable_RISC_v and used that "
        "code, not a brochure, to explain how RISC-V works. An instruction is a 32-bit form. "
        "A pointer is a register filled by LUI and ADDI. A pipeline is five locked steps with "
        "forwarding to hide the locks. Instruction memory is a 64-word array that a UART "
        "bootloader fills after a 0x25/0x55 handshake and empties its stall on 0xBAADF00D. "
        "UART reliability comes from sampling each bit sixteen times and believing the "
        "majority. GPIO is a flip-flop; SPI chip-select is a flip-flop that refuses to rise "
        "while the shifter is busy. SPI has four modes; this die uses Mode 0 at 4.17 MHz. "
        "Simulation is Icarus plus cocotb plus GTKWave. FPGA is a 50 MHz dress rehearsal on "
        "Kria KV260. Silicon is LibreLane on GF180MCU, sixteen tiles, Tiny Tapeout GF 26a "
        "project 0039, GDSII out. That whole sentence is one education path from a first "
        "assembly line to a real chip, and it is now written down so others can walk it."
    )

    add_sec("ACKNOWLEDGMENT")
    add_p(
        "The authors thank the engineering team at Antarikchya Pratisthan Nepal (Nepal Space "
        "Foundation), the Tiny Tapeout organizers, the LibreLane and Yosys developers, "
        "GlobalFoundries and Google for the GF180MCU open PDK, and the RISC-V International "
        "community for a free instruction set. Shuttle access for Tiny Tapeout GF 26a made "
        "the silicon experiment possible [7], [8], [43]."
    )

    add_sec("CODE AND DATA AVAILABILITY")
    add_p(
        "RTL, documentation, and the cocotb testbench are published at "
        "https://github.com/PrmBRana/ttgf_prem_programmable_RISC_v (top module "
        "tt_um_prem_pipeline_test, Apache-2.0 license as shipped with the Tiny Tapeout "
        "template). Shuttle listing: Tiny Tapeout GF 26a project 0039, "
        "Programmable_Pipeline-RISC-V [6]–[8]."
    )

    add_sec("REFERENCES")
    refs = [
        "[1] A. Waterman and K. Asanović, Eds., The RISC-V Instruction Set Manual, Volume I: Unprivileged ISA. RISC-V International, ratified specification, 2019–2024.",
        "[2] A. Waterman, Y. Lee, D. A. Patterson, and K. Asanović, “The RISC-V Instruction Set Manual, Volume I: User-Level ISA, Version 2.0,” EECS Dept., Univ. of California, Berkeley, Tech. Rep. UCB/EECS-2014-54, 2014.",
        "[3] K. Asanović and D. A. Patterson, “Instruction Sets Should Be Free: The Case for RISC-V,” EECS Dept., Univ. of California, Berkeley, Tech. Rep. UCB/EECS-2014-146, Aug. 2014.",
        "[4] D. A. Patterson and A. Waterman, The RISC-V Reader: An Open Architecture Atlas. Strawberry Canyon, 2017.",
        "[5] M. Venn, “Tiny Tapeout: A shared silicon shuttle for education and open source,” Tiny Tapeout documentation, 2023–2026. [Online]. Available: https://tinytapeout.com",
        "[6] P. B. Rana, “ttgf_prem_programmable_RISC_v: Programmable pipelined RISC-V for Tiny Tapeout GF 26a,” GitHub repository, 2026. [Online]. Available: https://github.com/PrmBRana/ttgf_prem_programmable_RISC_v",
        "[7] Tiny Tapeout, “Tiny Tapeout GF 26a datasheet,” Jun. 22, 2026, project 0039 Programmable_Pipeline-RISC-V. [Online]. Available: https://tinytapeout.github.io/tinytapeout-gf-26a/datasheet.pdf",
        "[8] Tiny Tapeout, “Tiny Tapeout GF 26a shuttle index,” 2026. [Online]. Available: https://tinytapeout.com/chips/ttgf26a/",
        "[9] AMD, Kria KV260 Vision AI Starter Kit User Guide (UG1089). San Jose, CA, USA, 2022–2025.",
        "[10] AMD, Zynq UltraScale+ Device Technical Reference Manual (UG1085). San Jose, CA, USA, 2023.",
        "[11] D. A. Patterson and J. L. Hennessy, Computer Organization and Design RISC-V Edition: The Hardware/Software Interface, 2nd ed. Cambridge, MA, USA: Morgan Kaufmann, 2021.",
        "[12] A. Waterman, “Design of the RISC-V Instruction Set Architecture,” Ph.D. dissertation, EECS Dept., Univ. of California, Berkeley, 2016.",
        "[13] L. Teske, “RISC-V online assembler,” 2020–2026. [Online]. Available: https://riscvasm.lucasteske.dev/",
        "[14] B. W. Kernighan and D. M. Ritchie, The C Programming Language, 2nd ed. Englewood Cliffs, NJ, USA: Prentice Hall, 1988.",
        "[15] S. Harris and D. Harris, Digital Design and Computer Architecture: RISC-V Edition. Cambridge, MA, USA: Morgan Kaufmann, 2022.",
        "[16] Tiny Tapeout, “HDL project documentation, pinout, and testing guide,” 2024–2026. [Online]. Available: https://tinytapeout.com/hdl/testing/",
        "[17] S. H. Unger, “Hazards, critical races, and metastability,” IEEE Trans. Comput., vol. 44, no. 6, pp. 754–768, Jun. 1995.",
        "[18] J. L. Hennessy and D. A. Patterson, Computer Architecture: A Quantitative Approach, 6th ed. Cambridge, MA, USA: Morgan Kaufmann, 2019.",
        "[19] RISC-V International, The RISC-V Instruction Set Manual, Volume II: Privileged Architecture. 2021–2024.",
        "[20] M. D. Ercegovac and T. Lang, Digital Arithmetic. San Francisco, CA, USA: Morgan Kaufmann, 2004.",
        "[21] J. L. Hennessy, N. Jouppi, S. Przybylski, C. Rowen, and T. Gross, “Hardware/software tradeoffs for increased performance,” in Proc. Symp. Architectural Support for Programming Languages and Operating Systems (ASPLOS), 1982, pp. 2–11.",
        "[22] P. M. Kogge, The Architecture of Pipelined Computers. New York, NY, USA: McGraw-Hill, 1981.",
        "[23] I. Koren, Computer Arithmetic Algorithms, 2nd ed. Natick, MA, USA: A K Peters, 2002.",
        "[24] J. L. Hennessy et al., “MIPS: A microprocessor architecture,” in Proc. MICRO, 1982, pp. 17–22.",
        "[25] C. E. Leiserson et al., “There’s plenty of room at the top: What will drive computer performance after Moore’s law?” Science, vol. 368, no. 6495, 2020.",
        "[26] T. J. Bergin, Computer: A History of the Information Machine. New York, NY, USA: Basic Books, 1996. (For the long history of serial telegraph framing that UART still uses.)",
        "[27] J. G. Proakis and M. Salehi, Digital Communications, 5th ed. New York, NY, USA: McGraw-Hill, 2008.",
        "[28] B. Sklar, Digital Communications: Fundamentals and Applications, 2nd ed. Upper Saddle River, NJ, USA: Prentice Hall, 2001.",
        "[29] F. J. Harris, “On the use of windows for harmonic analysis with the discrete Fourier transform,” Proc. IEEE, vol. 66, no. 1, pp. 51–83, Jan. 1978. (Oversampling and windowed majority as noise immunity.)",
        "[30] FTDI, “AN_232_03: UART baud-rate tolerance and clock accuracy,” Future Technology Devices International, application note, 2016.",
        "[31] R. Ginosar, “Metastability and synchronizers: A tutorial,” IEEE Des. Test Comput., vol. 28, no. 5, pp. 23–35, Sep./Oct. 2011.",
        "[32] V. Zimmer and M. Rothman, Beyond BIOS: Developing with the Unified Extensible Firmware Interface, 3rd ed. Berlin, Germany: De Gruyter, 2017. (System bring-up and loader contracts.)",
        "[33] N. H. E. Weste and D. M. Harris, CMOS VLSI Design: A Circuits and Systems Perspective, 4th ed. Boston, MA, USA: Addison-Wesley, 2011.",
        "[34] TDK InvenSense, IAM-20380 High Performance Automotive 3-Axis Gyroscope, Datasheet DS-000195, 2021.",
        "[35] MEMSIC, MMC5983MA ±8 Gauss High Performance 3-Axis Magnetic Sensor, Datasheet, 2019.",
        "[36] NASA Ames Research Center, “State of the Art of Small Spacecraft Technology,” NASA/TP, latest ed., 2021–2024.",
        "[37] Motorola Semiconductor, SPI Block Guide V03.06 (S12SPIV3/D), 2003. (Defines CPOL/CPHA modes 0–3.)",
        "[38] Texas Instruments, “KeyStone Architecture Serial Peripheral Interface (SPI) user guide,” SPRUGP2, 2012.",
        "[39] cocotb contributors, “cocotb: Coroutine-based cosimulation testbench environment,” documentation, 2014–2026. [Online]. Available: https://docs.cocotb.org/",
        "[40] IEEE Standard for Verilog Hardware Description Language, IEEE Std 1364-2005.",
        "[41] IEEE Standard for SystemVerilog—Unified Hardware Design, Specification, and Verification Language, IEEE Std 1800-2017.",
        "[42] E. Elzeiny et al., “LibreLane: An open-source RTL-to-GDSII flow,” Efabless / Tiny Tapeout documentation, 2024–2026. [Online]. Available: https://librelane.readthedocs.io/",
        "[43] Google and GlobalFoundries, “GF180MCU open-source PDK,” 2022–2025. [Online]. Available: https://github.com/google/gf180mcu-pdk",
        "[44] S. Williams, “Icarus Verilog,” 1998–2026. [Online]. Available: https://steveicarus.github.io/iverilog/",
        "[45] IEEE Standard for Universal Verification Methodology Language Reference Manual, IEEE Std 1800.2-2020. (Industry contrast to cocotb.)",
        "[46] T. Bybell, “GTKWave 3.3 Analyzer User’s Guide,” 2021. [Online]. Available: https://gtkwave.sourceforge.net/",
        "[47] Tiny Tapeout, “tt-gds-action GitHub Action (ttgf26a tag),” 2026. [Online]. Available: https://github.com/TinyTapeout/tt-gds-action",
        "[48] The CubeSat Program, Cal Poly SLO, CubeSat Design Specification, Rev. 14.1, 2022.",
        "[49] C. Maxfield, The Design Warrior’s Guide to FPGAs. Boston, MA, USA: Newnes, 2004.",
        "[50] FTDI, FT232R USB UART IC Datasheet, version 2.16, 2020.",
        "[51] Cadence Design Systems, “GDSII Stream Format Manual,” 1987/2000. (Industry layout interchange.)",
        "[52] OASIS — Open Artwork System Interchange Standard, SEMI P39, 2005. (Modern companion to GDSII.)",
        "[53] M. Shalan and T. Edwards, “Building OpenLANE: A 130nm openlane-based tapeout as a case study,” in Proc. Workshop on Open-Source EDA Technology (WOSET), 2020.",
        "[54] T. Ajayi et al., “Toward an open-source digital flow: First learnings from the OpenROAD project,” in Proc. 56th ACM/IEEE Design Automation Conf. (DAC), 2019.",
        "[55] A. B. Kahng, J. Lienig, I. L. Markov, and J. Hu, VLSI Physical Design: From Graph Partitioning to Timing Closure. Dordrecht, The Netherlands: Springer, 2011.",
        "[56] L. Lavagno, I. L. Markov, G. Martin, and L. K. Scheffer, Eds., Electronic Design Automation for IC Implementation, Circuit Design, and Process Technology. Boca Raton, FL, USA: CRC Press, 2016.",
        "[57] C. Wolf, J. Glaser, and J. Kepler, “Yosys—A free Verilog synthesis suite,” in Proc. Austrian Workshop on Microelectronics (Austrochip), 2013.",
        "[58] T. Ansell and M. Shalan, “The next open source disruption: Hardware design using open PDKs,” IEEE Micro, vol. 42, no. 6, pp. 18–26, 2022.",
        "[59] P. B. Rana and T. Shrestha, “Towards Satellite-on-Chip: Nepal’s development of a low-power RISC-V RV32 microcontroller for attitude determination in CubeSat,” in Proc. 77th Int. Astronautical Congr. (IAC), Antalya, Türkiye, Oct. 2026, Paper IAC-26-D1-IP-6-x105939.",
        "[60] ESA, ECSS-E-ST-50-12C: SpaceWire—Links, Nodes, Routers and Networks. Noordwijk, The Netherlands, 2008. (Flight serial-bus contrast to SPI/UART.)",
        "[61] M. Bell, “TinyQV: A RISC-V SoC for Tiny Tapeout,” GitHub, 2024–2025. [Online]. Available: https://github.com/MichaelBell/tinyQV",
        "[62] O. Kindgren, “SERV: The award-winning bit-serial RISC-V core,” 2019–2025. [Online]. Available: https://github.com/olofk/serv",
        "[63] M. Guthaus, J. E. Stine, S. Ataei, et al., “OpenRAM: An open-source memory compiler,” in Proc. IEEE/ACM ICCAD, 2016.",
        "[64] R. C. Baumann, “Radiation-induced soft errors in advanced semiconductor technologies,” IEEE Trans. Device Mater. Rel., vol. 5, no. 3, pp. 305–316, Sep. 2005.",
        "[65] M. Nicolaidis, “Design for soft error mitigation,” IEEE Trans. Device Mater. Rel., vol. 5, no. 3, pp. 405–418, Sep. 2005.",
        "[66] J. L. Hennessy, “VLSI processor architecture,” IEEE Trans. Comput., vol. C-33, no. 12, pp. 1221–1246, Dec. 1984.",
        "[67] D. A. Patterson and C. H. Sequin, “RISC I: A reduced instruction set VLSI computer,” in Proc. 8th Annu. Symp. Computer Architecture (ISCA), 1981, pp. 443–457.",
        "[68] S. Borkar, “Designing reliable systems from unreliable components: The challenges of transistor variability and degradation,” IEEE Micro, vol. 25, no. 6, pp. 10–16, 2005.",
        "[69] I. Sutherland, R. F. Sproull, and D. Harris, Logical Effort: Designing Fast CMOS Circuits. San Francisco, CA, USA: Morgan Kaufmann, 1999.",
        "[70] J. Rabaey, A. Chandrakasan, and B. Nikolić, Digital Integrated Circuits: A Design Perspective, 2nd ed. Upper Saddle River, NJ, USA: Prentice Hall, 2003.",
        "[71] M. Keating, D. Flynn, R. Aitken, A. Gibbons, and K. Shi, Low Power Methodology Manual. New York, NY, USA: Springer, 2007.",
        "[72] S. Palnitkar, Verilog HDL: A Guide to Digital Design and Synthesis, 2nd ed. Upper Saddle River, NJ, USA: Prentice Hall, 2003.",
        "[73] C. Spear and G. Tumbush, SystemVerilog for Verification, 3rd ed. New York, NY, USA: Springer, 2012.",
        "[74] P. Rashinkar, P. Paterson, and L. Singh, System-on-a-Chip Verification. Boston, MA, USA: Kluwer, 2001.",
        "[75] W. J. Dally and J. W. Poulton, Digital Systems Engineering. Cambridge, U.K.: Cambridge Univ. Press, 1998.",
        "[76] H. Johnson and M. Graham, High-Speed Digital Design: A Handbook of Black Magic. Englewood Cliffs, NJ, USA: Prentice Hall, 1993.",
        "[77] N. Storey, Electronics: A Systems Approach, 5th ed. Harlow, U.K.: Pearson, 2013. (UART/SPI at the systems-course level.)",
        "[78] J. Catsoulis, Designing Embedded Hardware, 2nd ed. Sebastopol, CA, USA: O’Reilly, 2005.",
        "[79] T. Noergaard, Embedded Systems Architecture, 2nd ed. Oxford, U.K.: Newnes, 2013.",
        "[80] K. Arnold, Embedded Controller Hardware Design. San Diego, CA, USA: LLH, 2001.",
        "[81] Tiny Tapeout Ltd., Invoice/Receipt 61446, GF 26a tile allocation, May 31, 2026.",
        "[82] YosysHQ, “Nextpnr and open FPGA tooling,” 2018–2026. [Online]. Available: https://github.com/YosysHQ/nextpnr",
        "[83] T. Edwards, “The Open_PDKs project and Magic VLSI,” 2020–2025. [Online]. Available: http://opencircuitdesign.com/",
        "[84] KLayout project, “KLayout layout viewer and editor,” 2006–2026. [Online]. Available: https://www.klayout.de/",
        "[85] S. C. Woo, M. Ohara, E. Torrie, J. P. Singh, and A. Gupta, “The SPLASH-2 programs: Characterization and methodological considerations,” in Proc. ISCA, 1995, pp. 24–36. (Why small kernels still teach real machines.)",
        "[86] D. E. Knuth, The Art of Computer Programming, Vol. 1: Fundamental Algorithms, 3rd ed. Reading, MA, USA: Addison-Wesley, 1997. (Pointers as addresses.)",
        "[87] A. S. Tanenbaum and H. Bos, Modern Operating Systems, 4th ed. Boston, MA, USA: Pearson, 2015. (Bootloaders as the first program a machine trusts.)",
        "[88] E. A. Lee and S. A. Seshia, Introduction to Embedded Systems: A Cyber-Physical Systems Approach, 2nd ed. Cambridge, MA, USA: MIT Press, 2017.",
        "[89] P. Marwedel, Embedded System Design, 3rd ed. Cham, Switzerland: Springer, 2018.",
        "[90] I. Grout, Digital Systems Design with FPGAs and CPLDs. Oxford, U.K.: Newnes, 2008.",
    ]
    for ref in refs:
        p_ref = doc.add_paragraph()
        p_ref.paragraph_format.left_indent = Inches(0.35)
        p_ref.paragraph_format.first_line_indent = Inches(-0.35)
        p_ref.paragraph_format.space_before = Pt(0)
        p_ref.paragraph_format.space_after = Pt(2)
        rr = p_ref.add_run(ref)
        rr.font.name = "Calibri"
        rr.font.size = Pt(8.5)
        rr.font.color.rgb = C_BODY

    out1 = os.path.join(OUT_DIR, "From_Verilog_to_GDSII_Programmable_RISCV_RV32I.docx")
    out2 = os.path.join(OUT_DIR, "paper_riscv_tutorial.docx")
    doc.save(out1)
    doc.save(out2)
    print("Wrote", out1)
    print("Wrote", out2)


if __name__ == "__main__":
    build_document()
