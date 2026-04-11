@app.post("/generate")
def generate_proposal(payload: GenerateRequest):
    prs = Presentation("templates/template_ninja.pptx")

    section = payload.sections[0]
    item = section.items[0]
    item_total = item.quantity * item.unit_price
    freight_value = section.freight_value or 0

    data = {
        "proposal_number": payload.proposal.proposal_number,
        "client_name": payload.proposal.client_name,
        "payment_method": payload.proposal.payment_method,
        "delivery_date": payload.proposal.delivery_date,
        "notes": payload.proposal.notes,
        "seller_name": "João Victor Ferrigno",
        "seller_phone": "(11) 99999-9999",
        "seller_email": "joao@empresa.com",
        "seller_description": "Teste vendedor",
        "item_name": item.item_name,
        "item_subtitle": item.item_subtitle,
        "item_index": str(item.item_index),
        "item_display_index": str(item.item_index + 1),
        "item_description": item.item_description,
        "item_code": item.item_code,
        "quantity": str(item.quantity),
        "unit_price": f"{item.unit_price:.2f}",
        "item_total": f"{item_total:.2f}",
        "section_total": f"{section.items[0].quantity * section.items[0].unit_price:.2f}",
        "freight": f"{freight_value:.2f}",
        "seller_image_url": "https://dummyimage.com/400x400/cccccc/000000.png&text=Seller",
        "item_image_url": "https://dummyimage.com/400x400/cccccc/000000.png&text=Item",
    }

    replace_text_placeholders_on_slide(prs.slides[8], data)
    replace_named_images_on_slide(prs.slides[8], data)

    replace_text_placeholders_on_slide(prs.slides[9], data)
    replace_named_images_on_slide(prs.slides[9], data)

    filename = f"proposta_teste_{str(uuid.uuid4())[:4]}.pptx"
    filepath = f"/tmp/{filename}"
    prs.save(filepath)

    return FileResponse(
        path=filepath,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
