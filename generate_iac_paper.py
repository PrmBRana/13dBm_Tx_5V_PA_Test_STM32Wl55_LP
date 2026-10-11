#!/usr/bin/env python3
"""
Single IAC 2026 manuscript: one strong research paper that unifies
the SatOC mission story with the taped-out RISC-V implementation.
Paper ID: IAC-26, D1,IP,6,x105939
Format: IAC 2026 guidelines (US Letter, Times New Roman 10 pt,
title/abstract full width, body two-column, recommended ≤ 15 pages).
"""

import os
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.shared import Inches, Pt, RGBColor, Twips, Emu

OUT_DIR = "/home/prem/Documents/research paper"
PAPER_ID = "IAC-26, D1,IP,6,x105939"
BLACK = RGBColor(0, 0, 0)
GRAY = RGBColor(60, 60, 60)


def set_run(run, size=10, bold=False, italic=False, color=BLACK, name="Times New Roman", center=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def add_page_field(paragraph, prefix, suffix):
    run = paragraph.add_run(prefix)
    set_run(run, size=8, italic=True, color=GRAY)
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), "begin")
    paragraph.add_run()._r.append(fld)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    paragraph.add_run()._r.append(instr)
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    paragraph.add_run()._r.append(fld2)
    run = paragraph.add_run(suffix)
    set_run(run, size=8, italic=True, color=GRAY)


def set_two_columns(section, space=360):
    sectPr = section._sectPr
    for child in list(sectPr):
        if child.tag == qn("w:cols"):
            sectPr.remove(child)
    cols = OxmlElement("w:cols")
    cols.set(qn("w:num"), "2")
    cols.set(qn("w:space"), str(space))
    cols.set(qn("w:equalWidth"), "1")
    sectPr.append(cols)


def set_cell_bg(cell, hex_color):
    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'))


def set_cell_m(cell, t=60, b=60, l=60, r=60):
    cell._tc.get_or_add_tcPr().append(parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{t}" w:type="dxa"/>'
        f'<w:bottom w:w="{b}" w:type="dxa"/>'
        f'<w:left w:w="{l}" w:type="dxa"/>'
        f'<w:right w:w="{r}" w:type="dxa"/>'
        f'</w:tcMar>'
    ))


def borders(table):
    table._tbl.tblPr.append(parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="888888"/>'
        f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="888888"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'</w:tblBorders>'
    ))


def build():
    doc = Document()
    sec0 = doc.sections[0]
    sec0.page_width = Inches(8.5)
    sec0.page_height = Inches(11)
    sec0.top_margin = Inches(0.7)
    sec0.bottom_margin = Inches(0.7)
    sec0.left_margin = Inches(0.7)
    sec0.right_margin = Inches(0.7)
    sec0.different_first_page_header_footer = False

    hp = sec0.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = hp.add_run("77th International Astronautical Congress (IAC 2026), Antalya, Türkiye, 5–9 October 2026.")
    set_run(r, size=8, italic=True)
    hp2 = sec0.header.add_paragraph()
    hp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = hp2.add_run("Copyright ©2026 by the International Astronautical Federation (IAF). All rights reserved.")
    set_run(r, size=8, italic=True)

    fp = sec0.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_field(fp, f"{PAPER_ID}                    Page ", "")

    def P(text, size=10, bold=False, italic=False, align="left", space_after=6, space_before=0, first=0):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.08
        p.paragraph_format.first_line_indent = Inches(first)
        p.alignment = {
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }.get(align, WD_ALIGN_PARAGRAPH.LEFT)
        run = p.add_run(text)
        set_run(run, size=size, bold=bold, italic=italic)
        return p

    def SEC(text):
        return P(text, size=11, bold=True, space_before=10, space_after=4)

    def SUB(text):
        return P(text, size=10, italic=True, bold=True, space_before=7, space_after=3)

    def JP(text):
        return P(text, align="justify", first=0.15, space_after=5)

    def EQ(text, n):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(text)
        set_run(r, size=10, italic=True)
        r = p.add_run(f"     ({n})")
        set_run(r, size=10)

    def FIG(name, caption, w=Inches(3.15)):
        path = name if os.path.isabs(name) else os.path.join(OUT_DIR, name)
        if os.path.isfile(path):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
            p.add_run().add_picture(path, width=w)
        P(caption, size=8, italic=True, align="center", space_after=6)

    def TBL(headers, rows, caption, widths):
        P(caption, size=8, italic=True, bold=True, space_before=6, space_after=3)
        table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        borders(table)
        for i, h in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.width = widths[i]
            set_cell_bg(cell, "1F2937")
            set_cell_m(cell)
            hp = cell.paragraphs[0]
            hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = hp.add_run(h)
            set_run(r, size=7.5, bold=True, color=RGBColor(255, 255, 255))
        for ri, row in enumerate(rows):
            for ci, val in enumerate(row):
                cell = table.rows[ri + 1].cells[ci]
                cell.width = widths[ci]
                set_cell_bg(cell, "F3F4F6" if ri % 2 else "FFFFFF")
                set_cell_m(cell)
                rp = cell.paragraphs[0]
                r = rp.add_run(str(val))
                set_run(r, size=7.5)
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    def CODE(text, caption):
        P(caption, size=8, italic=True, bold=True, space_before=4, space_after=2)
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.left_indent = Inches(0.08)
        r = p.add_run(text.strip())
        set_run(r, size=7.5, name="Courier New")

    # ========== TITLE BLOCK (full width) ==========
    P(PAPER_ID, size=11, bold=True, align="center", space_after=8)
    P(
        "Towards Satellite-on-Chip: Nepal’s Development of a Low-Power RISC-V RV32 "
        "Microcontroller for Attitude Determination in CubeSat",
        size=13, bold=True, align="center", space_after=8,
    )
    P("Prem Bahadur Rana a*, Trishna Shrestha b", size=11, align="center", space_after=2)
    P(
        "a Nepal Space Foundation / Antarikchya Pratisthan Nepal, Kathmandu, Nepal, prem@antarikchya.org.np\n"
        "b Nepal Space Foundation / Antarikchya Pratisthan Nepal, Kathmandu, Nepal, trishna@antarikchya.org.np\n"
        "* Corresponding Author",
        size=9, italic=True, align="center", space_after=8,
    )

    P("Abstract", size=11, bold=True, space_after=3)
    P(
        "CubeSats are constrained in electrical power, board area, mass and the number of chips that can be "
        "flown. Commercial microcontrollers are over-provisioned for a 1U bus and cannot be rewritten at the "
        "pipeline or peripheral-map level, which limits processor-level research in emerging space programmes "
        "such as Nepal. A Satellite-on-Chip (SatOC) that places the processor, memories, sensor buses and "
        "fault handling on one die would cut inter-chip wiring, volume and recurring cost. This paper reports "
        "the first indigenous step on that path: a programmable five-stage RISC-V RV32I microcontroller, "
        "written in Verilog, verified in assembly-level simulation and GTKWave, prototyped at 50 MHz on an "
        "AMD Kria KV260 FPGA, and taped out on sixteen tiles of the GlobalFoundries GF180MCU 180 nm "
        "open-source PDK through Tiny Tapeout GF 26a (project 0039, top module tt_um_prem_pipeline_test). "
        "After reset a UART bootloader waits for handshake 0x25, returns ACK 0x55, writes little-endian "
        "32-bit instructions into a 64-word instruction memory, and releases the pipeline on sentinel "
        "0xBAADF00D. Software names hardware with LUI/ADDI pointers and talks to UART2, GPIO and an SPI "
        "Mode-0 master through ordinary load/store addresses. The UART receiver oversamples each bit "
        "sixteen times and majority-votes the samples so firmware upload survives clock drift. GPIO2/GPIO3 "
        "hold chip-select until SPI is idle, matching IAM-20380 gyroscope and MMC5983MA magnetometer "
        "timing. The paper unifies the instruction-set explanation, the RTL, the laboratory FPGA path and "
        "the RTL-to-GDSII flow into one mission paper. The prototype is intended for in-orbit technology "
        "demonstration on NepaliSat-2.",
        size=10, align="justify", space_after=6,
    )
    P(
        "Keywords: CubeSat; RISC-V RV32I; Satellite-on-Chip; five-stage pipeline; UART 16× oversampling; "
        "SPI Mode 0; GPIO chip-select; hardware bootloader; FPGA; Tiny Tapeout; GF180MCU; GDSII; NepaliSat-2",
        size=9, italic=True, space_after=10,
    )

    # ========== BODY: two columns ==========
    sec1 = doc.add_section(WD_SECTION.CONTINUOUS)
    sec1.top_margin = Inches(0.7)
    sec1.bottom_margin = Inches(0.7)
    sec1.left_margin = Inches(0.7)
    sec1.right_margin = Inches(0.7)
    set_two_columns(sec1, 360)
    # Keep header/footer linked to the first section so the IAC banner is not duplicated.

    SEC("1. Introduction")
    JP(
        "CubeSats have become the standard low-cost platform for technology demonstration, science and "
        "education because launch and development costs are far below those of conventional satellites [1]. "
        "A 1U bus (10 cm × 10 cm × 10 cm, ≤ 1.33 kg) nevertheless leaves little orbit-average power "
        "(typically 1–2.5 W), little board area, and little tolerance for extra integrated circuits [1], [2]. "
        "The on-board computer (OBC) must schedule the mission, interrogate attitude sensors, packet "
        "telemetry and remain alive across launch vibration and thermal-vacuum cycling."
    )
    JP(
        "Commercial off-the-shelf (COTS) microcontrollers are used because their tool chains are mature. "
        "They are also a poor match to a national research goal. The silicon cannot be edited; unused USB, "
        "CAN or Ethernet blocks still leak power; and the processor, memories and sensor PHYs remain "
        "separate chips joined by board traces [2], [3]. For Nepal, which entered orbit with NepaliSat-1 "
        "in 2019 under the BIRDS programme [4], remaining a permanent buyer of closed avionics would "
        "freeze processor-level skill at the connector. The alternative is an open instruction set and an "
        "open foundry path."
    )
    JP(
        "RISC-V is a royalty-free ISA [5]–[7]. A laboratory may implement only the integer base (RV32I), "
        "choose a five-stage pipeline, map peripherals to ordinary load/store addresses, and send the "
        "same Verilog to an FPGA and then to a multi-project wafer. That is the SatOC programme of the "
        "Nepal Space Foundation: first a programmable microcontroller with the buses a CubeSat attitude "
        "suite actually uses, then, on later dice, SRAM, I2C, flash, watchdogs and radiation hardening [3], [8]."
    )
    JP(
        "This manuscript is the single IAC paper for that first step. The public RTL is "
        "https://github.com/PrmBRana/ttgf_prem_programmable_RISC_v (top module "
        "tt_um_prem_pipeline_test) [9]. Section 2 is the methodology: how that code is organised, "
        "how the RV32 pipeline fetches and executes instructions, how SPI/GPIO talk to attitude "
        "sensors, how the design is verified, how it is tested on FPGA, and how LibreLane turns "
        "the same Verilog into GDSII on Tiny Tapeout GF 26a. Sections 3–5 report results, limits "
        "and the SatOC roadmap."
    )

    SEC("2. Methodology")
    JP(
        "The method is to implement one Verilog microcontroller, prove it with assembly and "
        "waveforms, run it on an FPGA, then harden the identical HDL on an open 180 nm PDK. "
        "Every subsection below names the GitHub file that performs the work [9]."
    )

    SUB("2.1 System architecture and design approach")
    JP(
        "The design approach is mission-first, then hardware. A 1U attitude board needs a processor "
        "that can be programmed after reset, a serial link to the OBC, and a synchronous sensor bus "
        "with independent chip-selects. Interrupts, caches and unused PHYs are omitted so the core "
        "fits sixteen Tiny Tapeout tiles and remains deterministic [10], [17]. The same Verilog is "
        "the source for simulation, FPGA and ASIC; only clock-period parameters change."
    )
    JP(
        "Fig. 1 is the architecture in docs/img/block.png of the repository [9]. The Tiny Tapeout "
        "wrapper tt_um_prem_pipeline_test.v is the only module that sees the pads. It maps "
        "ui[0]/uo[0] to boot UART1, ui[1]/uo[1] to application UART2, and uio[0:5] to SPI CS, "
        "MOSI, MISO, SCLK, GPIO1 and a second CS, matching info.yaml. Inside the wrapper sits "
        "pipeline.v, which is the SoC: reset synchroniser, bootloader, instruction memory, "
        "five-stage CPU, MMIO decoder and peripherals. There are two operating modes. After "
        "reset, stall_pro holds the programme counter and UART1 fills instruction memory. After "
        "sentinel 0xBAADF00D the stall drops and the pipeline fetches from address 0."
    )
    FIG("block_diagram_repo.png",
        "Fig. 1. System architecture from the GitHub repository: host UART1 bootloader, 64-word "
        "instruction memory, five-stage RISC-V, UART2 / SPI / GPIO.")
    JP(
        "Peripherals are not special CPU instructions. DataMem.v compares the MEM-stage address "
        "against exact 32-bit constants (Table 1). A store to 0x4000_0000 starts an SPI byte; a "
        "load from 0x4000_000C reads RX-valid. This memory-mapped I/O is the design contract "
        "between assembly software and silicon [9], [16]."
    )
    TBL(
        ["Region", "Address", "RTL use"],
        [
            ["IMEM", "PC[7:2] → 64 words", "Instruction_Mem.v, boot write port"],
            ["UART2 TX/RX", "0x1000_0000 / 04", "uart0.v via DataMem.v"],
            ["UART2 status", "0x1000_0008 / 0C", "TX busy/pending; RX valid"],
            ["GPIO1 LED", "0x3000_0000", "gpio1_io in gpio.v"],
            ["GPIO2 CS", "0x3000_0004", "IAM-20380 CS_N hold-off"],
            ["GPIO3 CS", "0x3000_0008", "MMC5983MA CS_N hold-off"],
            ["SPI TX/RX", "0x4000_0000 / 08", "spi_master in spi.v"],
            ["SPI status", "0x4000_0004 / 0C", "busy, pending, RX valid"],
        ],
        "Table 1. Memory map implemented in DataMem.v [9].",
        [Inches(0.85), Inches(1.05), Inches(1.15)],
    )
    JP(
        "info.yaml declares clock_hz 25000000 and tiles 4x4. That is the architectural budget: "
        "25 MHz, about 1.36 mm × 0.64 mm of user area, 256 bytes of instruction memory, no "
        "general-purpose SRAM. Software therefore lives in registers and talks to the world "
        "only through the map in Table 1."
    )

    SUB("2.2 RV32 RISC-V processor architecture")
    JP(
        "A processor repeatedly reads a 32-bit word, decides what it means, and updates registers "
        "or pins. RISC-V is the public agreement on those words [5], [6]. This core implements "
        "RV32I: 32 registers x0–x31, a 32-bit PC, x0 hard-wired to zero [5], [13]. Every "
        "instruction is four bytes, so sequential fetch is PC ← PC + 4 (PC_incre.v). The lowest "
        "seven bits are the opcode. Fields rs1, rs2 and rd always sit in the same bit positions "
        "so the register file can start a read while Control_Unit.v is still classifying the "
        "instruction [5], [14]:"
    )
    EQ("instruction = {funct7, rs2, rs1, funct3, rd, opcode}   (R-type)", "1")
    JP(
        "Five shapes cover the ISA subset (Table 2). imm.v reconstructs I/S/B/J/U immediates. "
        "B-type and J-type always append a trailing zero because targets are even-byte aligned [5]. "
        "LUI and AUIPC need extra operand steering that many textbook cores miss: Control_Unit.v "
        "sets ALUSrcA = 10₂ so LUI adds the immediate to zero, and ALUSrcA = 01₂ so AUIPC adds "
        "it to PC_E [9]."
    )
    TBL(
        ["Type", "Examples", "RTL meaning"],
        [
            ["R", "add, and, or, xor, slt", "rd = f(rs1, rs2); ALUType 00"],
            ["I", "addi, lw, jalr, slli", "rd = f(rs1, ImmExt)"],
            ["S", "sw", "MemWrite; address = rs1 + S-imm"],
            ["B", "beq, bne, blt, bge", "ALUType 10; Zero is the condition"],
            ["U", "lui, auipc", "lui: pass B; auipc: PC + U-imm"],
            ["J", "jal / jalr", "ResultSrc = PC+4; Jump / JumpR"],
        ],
        "Table 2. RV32I formats decoded by Control_Unit.v.",
        [Inches(0.45), Inches(1.1), Inches(1.5)],
    )
    JP(
        "The microarchitecture is the classic five-stage pipeline in pipeline.v (Fig. 2) [13], [22]. "
        "IF: pc_register (clk.v) holds PCF; PCSelect_MUX chooses PCF+4 or PCTarget. IMEM is read "
        "combinationally at PCF[7:2]. ID: IF_ID_stage.v latches the word (or flushes it to a NOP "
        "on a taken branch). pipeline.v slices INSTR_op, INSTR_f3, INSTR_f7, rs1, rs2, rd. "
        "Reg_File.v is 32×32 with write-through bypass; x0 never writes. EX: EX_stage.v is the "
        "ID/EX latch. Forwarding muxes replace stale RD1/RD2 with ALUResultM or ResultW. The ALU "
        "adds, subtracts, logic-ops, shifts, or compares. JALR clears bit 0 of the target [5]:"
    )
    EQ("PCTarget = JumpR ? ((SrcA + Imm) & ~1) : (PC_E + Imm)", "2")
    EQ("PCSrcE   = (ZeroE ∧ BranchE) ∨ JumpE", "3")
    FIG("Screenshot from 2026-09-06 12-34-10.png",
        "Fig. 2. Five-stage RV32I datapath corresponding to pipeline.v, Hazard_Unit.v and ALU.v.")
    JP(
        "MEM: MEM_stage.v pipelines the ALU result into DataMem.v (MMIO only). WB: Write_back.v "
        "selects ALU result, load data or PC+4. Hazard_Unit.v forwards MEM over WB, stalls one "
        "clock on load-use, and flushes IF/ID on a taken branch. A pointer is just a register: "
        "lui x10, 0x40000 then addi x11, x10, 8 builds the SPI_RX address. Software never uses "
        "a dedicated I/O opcode [9], [15]."
    )
    EQ("pointer = (U-immediate << 12) + I-immediate", "4")

    SUB("2.3 Sensor interface and attitude monitoring")
    JP(
        "Attitude monitoring on a CubeSat needs a three-axis rate and a three-axis magnetic field "
        "[2], [18], [19]. This first die does not contain the sensors; it contains the bus that "
        "will clock them. GPIO2 (0x3000_0004) is documented in gpio.v as IAM-20380 CS_N. GPIO3 "
        "(0x3000_0008) is MMC5983MA CS_N. Both FSMs assert CS immediately on a store of 0 and "
        "refuse to de-assert until spi_busy and spi_pending are low, so the last SCLK edge is "
        "not cut off [9], [29]. GPIO1 is a status LED."
    )
    JP(
        "The sensor bus is SPI. spi.v is a parameterized master (DATA_WIDTH, CPOL, CPHA, CLK_DIV). "
        "Four Motorola modes are selected by polarity and phase (Table 3). The taped-out instance "
        "in pipeline.v is Mode 0, 8 bits, CLK_DIV = 3:"
    )
    EQ("f_SCLK = f_clk / (2 × CLK_DIV) = 25 MHz / 6 ≈ 4.17 MHz", "5")
    TBL(
        ["Mode", "CPOL", "CPHA", "Sample", "This silicon"],
        [
            ["0", "0", "0", "Rising", "Default in pipeline.v"],
            ["1", "0", "1", "Falling", "Available by parameter"],
            ["2", "1", "0", "Falling", "Available by parameter"],
            ["3", "1", "1", "Rising", "Available by parameter"],
        ],
        "Table 3. SPI modes implemented in spi.v [9], [29].",
        [Inches(0.45), Inches(0.42), Inches(0.42), Inches(0.65), Inches(1.1)],
    )
    JP(
        "Mode 0 drives MOSI before the first rising edge and samples MISO on that edge. A write "
        "to 0x4000_0000 arms a full-duplex byte; the received byte appears at 0x4000_0008 when "
        "RX-valid is set. There is no IRQ: firmware polls. UART2 (uart0.v) is the OBC link at "
        "the same 16× oversampled 115 200 baud class as the boot UART. I2C and flash, named as "
        "SatOC goals, are not in this GitHub tree; they are the next die once SRAM fits [3], [8]."
    )
    JP(
        "The attitude-monitoring sequence used in docs/info.md and test/test.py is: build LUI/ADDI "
        "pointers; assert GPIO2; poll SPI TX; store dummy 0xAA (many MEMS devices need clocks more "
        "than a command on the first beat); poll SPI RX; poll UART2 TX; store the byte to the OBC; "
        "repeat; release CS only after SPI is idle [9], [18]. A GPIO blink kernel with software "
        "delays proves branch timing independently of SPI."
    )
    CODE(
        "lui  x10, 0x40000     # SPI = 0x40000000\n"
        "lui  x18, 0x30000     # GPIO = 0x30000000\n"
        "lui  x19, 0x10000     # UART2 = 0x10000000\n"
        "sw   x0,  4(x18)      # CS low (gyro)\n"
        "# poll 0x40000004; sw 0xAA; poll 0x4000000C; lw RX\n"
        "# poll UART2 TX status; sw byte to 0x10000000\n"
        "sw   x3,  4(x18)      # CS high after idle\n"
        "ecall",
        "Listing 1. How the GitHub assembly talks to sensors and the OBC [9].",
    )

    SUB("2.4 RTL design and verification")
    JP(
        "All DUT files are synthesizable Verilog-2001 with `default_nettype none [9], [27], [49]. "
        "Table 4 is the source list from info.yaml. The UART (uart.v, uart0.v) is the reliability "
        "block. There is no shared clock on a serial line [25], [26]. Both UARTs tick at sixteen "
        "times the bit rate:"
    )
    EQ("BAUD_DIV = f_clk / (BaudRate × 16) = 25×10^6 / (115200 × 16) → 13", "6")
    JP(
        "The true ASIC baud is therefore 120 192 (+4.3 %). RX is two-flop synchronised [28]. "
        "Start and data bits are majority-voted over 16 ticks (threshold 8), so a glitch does not "
        "unlock the bootloader. uart_bootloader.v is the programmer: stall_pro = 1; byte 0x25 → "
        "ACK 0x55 (else NACK 0xFF); four little-endian bytes become a word; 0xBAADF00D ends "
        "writes and releases the pipeline (Table 5). Instruction_Mem.v is 64 × 32 bits with a "
        "boot write port and a combinational fetch port on PC[7:2]."
    )
    TBL(
        ["File", "Role in the GitHub RTL"],
        [
            ["tt_um_prem_pipeline_test.v", "Pad wrapper (only module with ui/uo/uio)"],
            ["pipeline.v", "SoC: CPU + boot + UART2 + SPI + GPIO"],
            ["Control_Unit.v / ALU.v", "ISA decode and integer/branch ALU"],
            ["Reg_File.v / imm.v", "x0–x31 and I/S/B/J/U immediates"],
            ["Hazard_Unit.v", "Forward, load-use stall, branch flush"],
            ["Instruction_Mem.v", "64-word programmable IMEM"],
            ["DataMem.v", "Exact-match MMIO decoder"],
            ["uart.v / uart_bootloader.v", "16× UART1 and 0x25 loader"],
            ["uart0.v / spi.v / gpio.v", "UART2, SPI master, CS/LED"],
        ],
        "Table 4. How the GitHub source files cooperate [9].",
        [Inches(1.25), Inches(1.95)],
    )
    TBL(
        ["Event", "Value", "Code action"],
        [
            ["Handshake", "0x25", "handshake_done; TX ACK"],
            ["ACK / NACK", "0x55 / 0xFF", "Host streams or retries"],
            ["Word", "4 bytes LE", "mem_we; addr_count++"],
            ["Sentinel", "0xBAADF00D", "boot_done; stall_pro → 0"],
        ],
        "Table 5. Bootloader protocol in uart_bootloader.v.",
        [Inches(0.75), Inches(0.95), Inches(1.35)],
    )
    JP(
        "Verification is Icarus Verilog + cocotb [30], [31]. test/tb.v instantiates the wrapper "
        "and dumps tb.fst. test/test.py (make -B) (1) starts a 40 ns / 25 MHz clock, (2) sends "
        "0x25 and expects 0x55, (3) uploads little-endian instruction words plus 0xBAADF00D, "
        "(4) runs a Mode-0 SPI slave that preloads MISO before the first rising edge, (5) logs "
        "UART2. GTKWave is opened with gtkwave tb.fst tb.gtkw [32] (Fig. 3). A correct trace "
        "shows stall_pro high during upload, we every four RX bytes, PCF stepping by 4 after "
        "the sentinel, and eight Mode-0 SCLK edges per byte while CS is low. Gate-level "
        "simulation repeats the same test on the GF180 netlist (make GATES=yes) [10]."
    )
    FIG("Screenshot from 2026-03-19 17-50-35.png",
        "Fig. 3. GTKWave clock check from the simulation flow in test/ [9], [32].")

    SUB("2.5 FPGA implementation and testing")
    JP(
        "An FPGA is the dress rehearsal before mask cost [20], [33]. The research programme mapped "
        "the UART and bootloader onto the programmable logic of an AMD Kria KV260 (Zynq "
        "UltraScale+ XCZU5EV) at 50 MHz from PL_CLK0, with an FTDI FT232R on a PMOD header "
        "[20], [21], [24] (Fig. 4–6). The test is the same protocol as silicon, only faster to "
        "iterate: bitstream, rst_n low then high, send 0x25, expect 0x55, stream words, send "
        "0xBAADF00D, watch GPIO or UART2. If 0xFF returns, the image is wrong or the line is "
        "inverted. The uploader’s auto-baud probe exists because a 50 MHz FPGA divider is not "
        "the 25 MHz ASIC divider; parameters in uart.v must match the board clock."
    )
    FIG("1789972494070.jpeg",
        "Fig. 4. Vivado IP Integrator twin of uart.v + uart_bootloader.v on the KV260.")
    FIG("Screenshot from 2026-09-29 11-54-13.png",
        "Fig. 5. Host UART Terminal & Instruction Uploader exercising the GitHub handshake.")
    FIG("1789972494046.jpeg",
        "Fig. 6. FPGA bench: Kria KV260, FTDI USB–UART, host laptop.")
    JP(
        "Laboratory findings that later entered the ASIC RTL include registered reset de-assert "
        "in pipeline.v, SPI busy/pending fan-out buffers, and CS hold-off in gpio2_io/gpio3_io "
        "[9], [17]. The repository also contains .github/workflows/fpga.yaml, which builds an "
        "ICE40UP5K bitstream for the official Tiny Tapeout FPGA playground from the same HDL [10]."
    )

    SUB("2.6 ASIC design flow and Tiny Tapeout")
    JP(
        "GDSII is the polygon file a mask shop uses [34]. Tiny Tapeout runs LibreLane on GitHub "
        "Actions so a student repository becomes a slot on a shared GF180MCU wafer [10]–[12], "
        "[35]–[39]. This project’s flow is literally the files in the repo. info.yaml sets "
        "top_module tt_um_prem_pipeline_test, tiles 4x4, clock_hz 25000000. src/config.json sets "
        "CLOCK_PERIOD 40 ns (25 MHz), PL_TARGET_DENSITY_PCT 60, hold-slack margins 0.1 / 0.05 ns, "
        "CLOCK_PORT clk, RUN_CTS 1. .github/workflows/gds.yaml calls "
        "TinyTapeout/tt-gds-action@ttgf26a with pdk: gf180mcuD, then precheck, gate-level test "
        "and a GDS viewer [9], [10]."
    )
    JP(
        "LibreLane, in human order, is: (1) Yosys reads the Verilog list and maps GF180 standard "
        "cells — inferred latches here would be a bug [35]; (2) floorplan places the TT pin list; "
        "(3) global placement at 60 % density; (4) clock-tree synthesis so every flop sees clk "
        "[17]; (5) routing on the 7-track 5 V library; (6) setup/hold repair against 40 ns; "
        "(7) Magic/KLayout/Netgen LVS and DRC [12], [37]; (8) write GDSII/LEF/DEF/SPEF and the "
        "gate-level netlist used by gl_test. After shuttle merge the board mux selects project "
        "0039, Programmable_Pipeline-RISC-V [11]. Silicon bring-up is the FPGA script at 25 MHz. "
        "Invoice 61446 records the 16-tile allocation [40]."
    )
    TBL(
        ["Step", "GitHub / tool evidence"],
        [
            ["RTL", "src/*.v listed in info.yaml"],
            ["Sim", "test/Makefile, test.py, tb.v, GTKWave"],
            ["FPGA", "workflows/fpga.yaml; KV260 laboratory"],
            ["Synth–GDS", "config.json + workflows/gds.yaml → LibreLane"],
            ["Sign-off", "precheck + GL test on gf180mcuD cells"],
            ["Shuttle", "TTGF26a address 39, 4×4 tiles"],
        ],
        "Table 6. End-to-end method from the public repository to silicon [9]–[12].",
        [Inches(0.85), Inches(2.35)],
    )

    SEC("3. Results")
    JP(
        "Table 7 compares the laboratory FPGA twin with the taped-out instance. Cocotb demonstrated "
        "handshake, little-endian upload, sentinel release, Mode-0 SPI and UART2 relay [9]. FPGA "
        "demonstrated the same protocol on physical pins (Fig. 5–6). ASIC timing is closed at "
        "25 MHz by CLOCK_PERIOD. The result that matters for IAC is not that 256 bytes of IMEM "
        "replace a commercial ADCS computer. It is that the GitHub tree is a closed, repeatable "
        "path from assembly to GDSII, sized for a NepaliSat-2 technology demonstration [4], [8], [11]."
    )
    TBL(
        ["Item", "FPGA (KV260)", "ASIC (TTGF26a / 0039)"],
        [
            ["Clock", "50 MHz", "25 MHz"],
            ["IMEM", "Block RAM", "64 × 32-bit DFF array"],
            ["Data RAM", "Available in PL", "None (MMIO only)"],
            ["UART", "16×, up to 480 kbaud", "16×, 115.2 kbaud target"],
            ["SPI", "Mode 0 lab slave", "Mode 0, ≈ 4.17 MHz"],
            ["GPIO", "PMOD / LED", "LED + 2 CS hold-off"],
            ["I2C / flash", "Mission software target", "Not in this GitHub RTL"],
            ["Path", "Vivado PL bitstream", "LibreLane GDSII, 4×4 tiles"],
        ],
        "Table 7. What the method produced on FPGA versus GF180 silicon.",
        [Inches(0.7), Inches(1.2), Inches(1.15)],
    )

    SEC("4. Discussion")
    JP(
        "Two cultures meet in this repository. One is computer architecture: Patterson’s pipeline, "
        "RISC-V immediates, forwarding [5], [13]. The other is spacecraft bring-up: FTDI cables, "
        "CS hold-off, majority-voted start bits, shuttle density knobs [10], [26], [29]. Compared "
        "with other Tiny Tapeout RISC-V cores [41], [42], this design is not the smallest. It is "
        "programmable without JTAG and it exposes the first buses a CubeSat attitude board needs: "
        "UART to the OBC, SPI to sensors, GPIO for CS and a beacon LED."
    )
    JP(
        "Limits are part of the method. There is no stack SRAM, no I2C master, no ECC and no TMR "
        "[43], [44]. The integer baud divider is coarse at 25 MHz. GF180 standard cells are not "
        "rad-hard. Before NepaliSat-2 flight use, packaged parts must see thermal-vacuum and "
        "random-vibration screening [2], [8]. The next SatOC die should add OpenRAM [45], a "
        "fractional UART divider, I2C, a watchdog, and parity or lockstep on the PC [3], [43]. "
        "If that sequence holds, future Nepali buses can stop adapting software to a foreign chip "
        "and start adapting the chip to the mission [3], [8]."
    )

    SEC("5. Conclusion")
    JP(
        "This paper reported one methodology, traced through one public repository. A five-stage "
        "RISC-V RV32I microcontroller was specified in Verilog, programmed by a 0x25/0x55/"
        "0xBAADF00D UART bootloader into a 64-word instruction memory, and operated through "
        "LUI/ADDI pointers to a 16×-oversampled UART, GPIO chip-selects and an SPI Mode-0 master "
        "at 4.17 MHz. Cocotb/GTKWave simulation and a 50 MHz Kria KV260 prototype preceded "
        "LibreLane tapeout on Tiny Tapeout GF 26a project 0039. The work is the first indigenous "
        "microcontroller step toward a Satellite-on-Chip for Nepal and is intended for in-orbit "
        "demonstration on NepaliSat-2."
    )

    SEC("Acknowledgements")
    JP(
        "The authors thank the engineering team at Antarikchya Pratisthan Nepal (Nepal Space "
        "Foundation), the Tiny Tapeout organisers, the LibreLane and Yosys developers, "
        "GlobalFoundries and Google for the GF180MCU open PDK, and RISC-V International for a "
        "free instruction set. Shuttle access on Tiny Tapeout GF 26a made the silicon experiment "
        "possible [10], [11], [39], [40]."
    )

    SEC("References")
    refs = [
        "[1] The CubeSat Program, Cal Poly San Luis Obispo, CubeSat Design Specification, Rev. 14.1, 2022.",
        "[2] NASA Ames Research Center, State of the Art of Small Spacecraft Technology, NASA/TP, 2021–2024.",
        "[3] D.A. Patterson and J.L. Hennessy, Computer Organization and Design RISC-V Edition, 2nd ed., Morgan Kaufmann, 2021.",
        "[4] BIRDS Project / Kyushu Institute of Technology, NepaliSat-1 mission documentation, 2019.",
        "[5] A. Waterman and K. Asanović (Eds.), The RISC-V Instruction Set Manual, Vol. I: Unprivileged ISA, RISC-V International.",
        "[6] K. Asanović and D.A. Patterson, “Instruction Sets Should Be Free: The Case for RISC-V,” UCB/EECS-2014-146, 2014.",
        "[7] A. Waterman, “Design of the RISC-V Instruction Set Architecture,” Ph.D. dissertation, UC Berkeley, 2016.",
        "[8] ESA, ECSS-E-ST-10-03C: Testing, Noordwijk, 2012; and ECSS-Q-ST-60C: Electrical, electronic and electromechanical components.",
        "[9] P.B. Rana, ttgf_prem_programmable_RISC_v, GitHub, 2026. https://github.com/PrmBRana/ttgf_prem_programmable_RISC_v",
        "[10] Tiny Tapeout, HDL documentation, pinout and testing guide, 2024–2026. https://tinytapeout.com",
        "[11] Tiny Tapeout, GF 26a datasheet and shuttle index, project 0039 Programmable_Pipeline-RISC-V, 22 June 2026.",
        "[12] LibreLane documentation, Efabless / Tiny Tapeout, 2024–2026. https://librelane.readthedocs.io/",
        "[13] S. Harris and D. Harris, Digital Design and Computer Architecture: RISC-V Edition, Morgan Kaufmann, 2022.",
        "[14] J.L. Hennessy and D.A. Patterson, Computer Architecture: A Quantitative Approach, 6th ed., Morgan Kaufmann, 2019.",
        "[15] L. Teske, RISC-V online assembler, https://riscvasm.lucasteske.dev/",
        "[16] E.A. Lee and S.A. Seshia, Introduction to Embedded Systems, 2nd ed., MIT Press, 2017.",
        "[17] N.H.E. Weste and D.M. Harris, CMOS VLSI Design, 4th ed., Addison-Wesley, 2011.",
        "[18] TDK InvenSense, IAM-20380 High Performance Automotive 3-Axis Gyroscope, DS-000195, 2021.",
        "[19] MEMSIC, MMC5983MA ±8 Gauss 3-Axis Magnetic Sensor, datasheet, 2019.",
        "[20] AMD, Kria KV260 Vision AI Starter Kit User Guide (UG1089), 2022–2025.",
        "[21] AMD, Zynq UltraScale+ Device Technical Reference Manual (UG1085), 2023.",
        "[22] P.M. Kogge, The Architecture of Pipelined Computers, McGraw-Hill, 1981.",
        "[23] A.S. Tanenbaum and H. Bos, Modern Operating Systems, 4th ed., Pearson, 2015.",
        "[24] FTDI, FT232R USB UART IC datasheet, v2.16, 2020; AN on UART baud tolerance, 2016.",
        "[25] J.G. Proakis and M. Salehi, Digital Communications, 5th ed., McGraw-Hill, 2008.",
        "[26] B. Sklar, Digital Communications: Fundamentals and Applications, 2nd ed., Prentice Hall, 2001.",
        "[27] IEEE Std 1364-2005, Verilog Hardware Description Language.",
        "[28] R. Ginosar, “Metastability and synchronizers: A tutorial,” IEEE Des. Test Comput., 28(5), 2011.",
        "[29] Motorola, SPI Block Guide V03.06 (S12SPIV3/D), 2003.",
        "[30] cocotb documentation, https://docs.cocotb.org/",
        "[31] S. Williams, Icarus Verilog, https://steveicarus.github.io/iverilog/",
        "[32] T. Bybell, GTKWave Analyzer User’s Guide, 2021.",
        "[33] C. Maxfield, The Design Warrior’s Guide to FPGAs, Newnes, 2004.",
        "[34] Cadence, GDSII Stream Format Manual, 1987/2000.",
        "[35] C. Wolf et al., “Yosys—A free Verilog synthesis suite,” Austrochip, 2013.",
        "[36] T. Ajayi et al., “Toward an open-source digital flow,” Proc. DAC, 2019.",
        "[37] M. Shalan and T. Edwards, “Building OpenLANE,” WOSET, 2020.",
        "[38] A.B. Kahng et al., VLSI Physical Design: From Graph Partitioning to Timing Closure, Springer, 2011.",
        "[39] T. Ansell and M. Shalan, “The next open source disruption: Hardware design using open PDKs,” IEEE Micro, 42(6), 2022.",
        "[40] Tiny Tapeout Ltd., Invoice/Receipt 61446, GF 26a 16-tile allocation, 31 May 2026.",
        "[41] M. Bell, TinyQV RISC-V SoC for Tiny Tapeout, 2024–2025.",
        "[42] O. Kindgren, SERV bit-serial RISC-V core, 2019–2025.",
        "[43] R.C. Baumann, “Radiation-induced soft errors in advanced semiconductor technologies,” IEEE Trans. Device Mater. Rel., 5(3), 2005.",
        "[44] M. Nicolaidis, “Design for soft error mitigation,” IEEE Trans. Device Mater. Rel., 5(3), 2005.",
        "[45] M. Guthaus et al., “OpenRAM: An open-source memory compiler,” Proc. ICCAD, 2016.",
        "[46] Google and GlobalFoundries, GF180MCU open-source PDK, 2022–2025.",
        "[47] D.A. Patterson and C.H. Sequin, “RISC I: A reduced instruction set VLSI computer,” Proc. ISCA, 1981.",
        "[48] J. Rabaey, A. Chandrakasan and B. Nikolić, Digital Integrated Circuits, 2nd ed., Prentice Hall, 2003.",
        "[49] S. Palnitkar, Verilog HDL, 2nd ed., Prentice Hall, 2003.",
        "[50] IAF, IAC 2026 Manuscript Guidelines, 2026.",
    ]
    for ref in refs:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.first_line_indent = Inches(-0.2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.05
        r = p.add_run(ref)
        set_run(r, size=8)

    out_docx = os.path.join(OUT_DIR, "IAC-26-D1-IP-6-x105939_Towards_Satellite_on_Chip_RISCV.docx")
    doc.save(out_docx)
    print("Wrote", out_docx, os.path.getsize(out_docx))
    return out_docx


if __name__ == "__main__":
    build()
