import argparse
import os

try:
    from pypdf import PdfReader, PdfWriter, Transformation
except ImportError:
    from PyPDF2 import PdfReader, PdfWriter, Transformation

MM_TO_PT = 72 / 25.4  # 1 mm in punten

PAPER_SIZES_MM = {
    "A4":  (210, 297),
    "A3":  (297, 420),
    "SRA4": (225, 320),
    "SRA3": (320, 450),
}

# Correctie voor duplex-afwijking van jouw printer
# positief Y = achterkant omhoog t.o.v. voorkant
BACK_OFFSET_X_MM = -2.5
BACK_OFFSET_Y_MM = 0.0


def mm_to_points(w_mm, h_mm):
    return w_mm * MM_TO_PT, h_mm * MM_TO_PT


def get_sheet_size(format_name):
    fmt = format_name.upper()
    if fmt not in PAPER_SIZES_MM:
        raise ValueError(
            f"Onbekend papierformaat '{format_name}'. Kies uit: A4, A3, SRA4, SRA3."
        )
    w_mm, h_mm = PAPER_SIZES_MM[fmt]
    return mm_to_points(w_mm, h_mm)


def compute_grid(card_w, card_h, usable_w, usable_h):
    """
    usable_w/h: bruikbaar gebied (na aftrek minimale printmarge).
    Berekent hoeveel kaarten er passen en hoeveel 'lege' marge
    binnen dat bruikbare gebied overblijft, zodat het grid
    daarbinnen gecentreerd staat.
    """
    cols = int(usable_w // card_w)
    rows = int(usable_h // card_h)

    if cols < 1 or rows < 1:
        raise ValueError(
            "De kaart is groter dan het bruikbare deel van het papierformaat – "
            "er past niet eens één kaart op het vel."
        )

    used_w = cols * card_w
    used_h = rows * card_h

    margin_x = (usable_w - used_w) / 2.0
    margin_y = (usable_h - used_h) / 2.0

    return cols, rows, margin_x, margin_y


def impose_side(
    writer,
    base_page,
    sheet_w,
    sheet_h,
    cols,
    rows,
    margin_x,
    margin_y,
    card_w_eff,
    card_h_eff,
    rotate=False,
):
    """
    Maakt één nieuw vel en zet daar base_page in een grid.
    'margin_x' en 'margin_y' zijn de afstanden van de papierrand
    tot het eerste kaartje (in punten).
    card_w_eff/h_eff zijn de effectieve kaartafmetingen in het grid
    (dus eventueel gewisseld bij rotatie).
    """
    new_page = writer.add_blank_page(width=sheet_w, height=sheet_h)

    card_w_raw = float(base_page.mediabox.width)
    card_h_raw = float(base_page.mediabox.height)

    for row in range(rows):
        for col in range(cols):
            x = margin_x + col * card_w_eff
            y = margin_y + row * card_h_eff

            if rotate:
                # 90° CCW om (0,0), daarna verplaatsen
                t = Transformation().rotate(90).translate(
                    tx=x + card_h_raw,
                    ty=y,
                )
            else:
                t = Transformation().translate(tx=x, ty=y)

            new_page.merge_transformed_page(base_page, t)

    return new_page


def crop_page_all_sides(page, trim_loss_mm):
    """Knipt aan alle kanten trim_loss_mm af van de mediabox."""
    if trim_loss_mm <= 0:
        return

    trim_pts = trim_loss_mm * MM_TO_PT
    llx, lly = page.mediabox.lower_left
    urx, ury = page.mediabox.upper_right

    new_llx = llx + trim_pts
    new_lly = lly + trim_pts
    new_urx = urx - trim_pts
    new_ury = ury - trim_pts

    if new_urx <= new_llx or new_ury <= new_lly:
        raise ValueError(
            "Te veel trim_loss_mm: er blijft geen bruikbaar kaartformaat over."
        )

    page.mediabox.lower_left = (new_llx, new_lly)
    page.mediabox.upper_right = (new_urx, new_ury)


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Step-and-repeat imposition voor 2-zijdige visitekaartjes "
            "(p1 = voor, p2 = achter). "
            "Houdt rekening met printmarges rondom, probeert automatisch 2 mm trim en rotatie, "
            "en corrigeert de achterkant met een vaste offset."
        )
    )
    parser.add_argument(
        "input_pdf",
        help="Input-PDF met minstens 2 pagina's (voorzijde = p1, achterzijde = p2).",
    )
    parser.add_argument(
        "output_pdf",
        nargs="?",
        default=None,
        help=(
            "(Optioneel) Naam van de gegenereerde impositie-PDF. "
            "Standaard: zelfde naam als input + '_PRINT'."
        ),
    )
    parser.add_argument(
        "--paper",
        "-p",
        default="A4",  # standaard A4
        help="Papierformaat: A4, A3, SRA4 of SRA3 (default: A4).",
    )
    parser.add_argument(
        "--margin-mm",
        type=float,
        default=None,
        help=(
            "Symmetrische minimale printmarge rondom (mm). "
            "Indien gezet, overschrijft margin-x-mm en margin-y-mm."
        ),
    )
    parser.add_argument(
        "--margin-x-mm",
        type=float,
        default=5.0,
        help="Minimale horizontale printmarge links en rechts (mm) (default: 5).",
    )
    parser.add_argument(
        "--margin-y-mm",
        type=float,
        default=5.0,
        help="Minimale verticale printmarge boven en onder (mm) (default: 5).",
    )

    args = parser.parse_args()

    # Outputnaam: zelfde als input + _PRINT, tenzij expliciet opgegeven
    if args.output_pdf is None:
        root, ext = os.path.splitext(args.input_pdf)
        if not ext:
            ext = ".pdf"
        output_pdf = f"{root}_PRINT{ext}"
    else:
        output_pdf = args.output_pdf

    # Symmetrische/rondom printmarges
    if args.margin_mm is not None:
        margin_x_mm = args.margin_mm
        margin_y_mm = args.margin_mm
    else:
        margin_x_mm = args.margin_x_mm
        margin_y_mm = args.margin_y_mm

    # Lees input
    reader = PdfReader(args.input_pdf)
    if len(reader.pages) < 2:
        raise ValueError(
            "De input-PDF moet minstens 2 pagina's hebben (voor- en achterzijde)."
        )

    front_page = reader.pages[0]
    back_page = reader.pages[1]

    # Originele kaartmaat (zonder trim)
    orig_w = float(front_page.mediabox.width)
    orig_h = float(front_page.mediabox.height)

    # Velformaat
    sheet_w, sheet_h = get_sheet_size(args.paper)

    # Minimale printmarges → bruikbaar gebied
    margin_x_min_pts = margin_x_mm * MM_TO_PT
    margin_y_min_pts = margin_y_mm * MM_TO_PT

    usable_w = sheet_w - 2 * margin_x_min_pts
    usable_h = sheet_h - 2 * margin_y_min_pts

    if usable_w <= 0 or usable_h <= 0:
        raise ValueError("Printmarge is te groot voor het gekozen papierformaat.")

    # Scenario's: (trim_mm = 0 of 2) x (rotate = False of True)
    scenarios = []
    trim_candidates_mm = [0.0, 2.0]

    for trim_mm in trim_candidates_mm:
        trim_pts = trim_mm * MM_TO_PT
        card_w = orig_w - 2 * trim_pts if trim_mm > 0 else orig_w
        card_h = orig_h - 2 * trim_pts if trim_mm > 0 else orig_h

        if card_w <= 0 or card_h <= 0:
            continue

        for rotate in (False, True):
            if rotate:
                card_w_eff = card_h
                card_h_eff = card_w
            else:
                card_w_eff = card_w
                card_h_eff = card_h

            try:
                cols, rows, inner_mx, inner_my = compute_grid(
                    card_w_eff,
                    card_h_eff,
                    usable_w,
                    usable_h,
                )
                capacity = cols * rows
            except ValueError:
                continue

            scenarios.append(
                {
                    "trim_mm": trim_mm,
                    "rotate": rotate,
                    "cols": cols,
                    "rows": rows,
                    "inner_mx": inner_mx,
                    "inner_my": inner_my,
                    "card_w_eff": card_w_eff,
                    "card_h_eff": card_h_eff,
                    "capacity": capacity,
                }
            )

    if not scenarios:
        raise ValueError(
            "Er past geen kaart op het gekozen papierformaat met de opgegeven printmarge, "
            "zelfs niet met 2 mm trim of rotatie."
        )

    # Kies scenario met maximale capaciteit.
    # Bij gelijke capaciteit: voorkeur voor minder trim en geen rotatie.
    scenarios.sort(
        key=lambda s: (s["capacity"], -s["trim_mm"], not s["rotate"]),
        reverse=True,
    )
    best = scenarios[0]

    # Trim nu daadwerkelijk toepassen (indien gekozen)
    if best["trim_mm"] > 0:
        crop_page_all_sides(front_page, best["trim_mm"])
        crop_page_all_sides(back_page, best["trim_mm"])

    # Afmetingen na eventuele trim
    final_card_w = float(front_page.mediabox.width)
    final_card_h = float(front_page.mediabox.height)

    # Bepaal effectieve kaartmaat in grid voor gekozen rotatie
    if best["rotate"]:
        card_w_eff = final_card_h
        card_h_eff = final_card_w
    else:
        card_w_eff = final_card_w
        card_h_eff = final_card_h

    # Grid-omvang en centrering op HET HELE VEL
    grid_w = best["cols"] * card_w_eff
    grid_h = best["rows"] * card_h_eff

    margin_x = (sheet_w - grid_w) / 2.0
    margin_y = (sheet_h - grid_h) / 2.0

    # Veiligheidscheck: margins mogen nooit kleiner zijn dan de minimale printmarges
    if margin_x < margin_x_min_pts - 0.01 or margin_y < margin_y_min_pts - 0.01:
        raise RuntimeError(
            "Interne fout: berekende marge kleiner dan minimale printmarge. "
            "Controleer de berekening."
        )

    print(f"Papier: {args.paper.upper()} ({sheet_w:.2f} x {sheet_h:.2f} pt)")
    print(
        f"Printmarge (min.): {margin_x_mm} mm links/rechts, "
        f"{margin_y_mm} mm boven/onder"
    )
    print("Gekozen scenario:")
    print(f"  - capaciteit : {best['capacity']} kaarten per vel")
    print(f"  - trim       : {best['trim_mm']} mm rondom")
    print(f"  - rotatie    : {'JA (90°)' if best['rotate'] else 'NEE'}")
    print(f"  - grid       : {best['cols']} kolommen x {best['rows']} rijen")
    print(
        "Eind-kaartformaat (zonder rekening te houden met rotatie): "
        f"{final_card_w:.2f} x {final_card_h:.2f} pt"
    )
    print(
        f"  - BACK_OFFSET_X_MM = {BACK_OFFSET_X_MM}, "
        f"BACK_OFFSET_Y_MM = {BACK_OFFSET_Y_MM}"
    )

    writer = PdfWriter()

    # Offset voor achterkant in punten
    back_offset_x = BACK_OFFSET_X_MM * MM_TO_PT
    back_offset_y = BACK_OFFSET_Y_MM * MM_TO_PT

    # Voorzijde
    impose_side(
        writer,
        front_page,
        sheet_w,
        sheet_h,
        best["cols"],
        best["rows"],
        margin_x,
        margin_y,
        card_w_eff,
        card_h_eff,
        rotate=best["rotate"],
    )

    # Achterzijde – zelfde grid, maar met kleine correctie voor duplex-afwijking
    impose_side(
        writer,
        back_page,
        sheet_w,
        sheet_h,
        best["cols"],
        best["rows"],
        margin_x + back_offset_x,
        margin_y + back_offset_y,
        card_w_eff,
        card_h_eff,
        rotate=best["rotate"],
    )

    with open(output_pdf, "wb") as f_out:
        writer.write(f_out)

    print(f"Gereed. Opgeslagen als: {output_pdf}")


if __name__ == "__main__":
    main()
