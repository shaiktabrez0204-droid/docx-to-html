from core.model import Image

# Test dict-like access
img = Image(image_id='test', relationship_id='rId1', source_path='word/media/img.png')

# Test __getitem__
try:
    val = img['anchor_paragraph_text']
    print('SUCCESS: img["anchor_paragraph_text"] =', val)
except TypeError as e:
    print('TypeError:', e)
except Exception as e:
    print('Other error:', e)

# Test __contains__ via __dict__
try:
    result = 'anchor_paragraph_text' in img.__dict__
    print('SUCCESS: "anchor_paragraph_text" in img.__dict__:', result)
except Exception as e:
    print('Error:', e)

# Test __contains__ via hasattr
try:
    result = hasattr(img, 'anchor_paragraph_text')
    print('SUCCESS: hasattr(img, "anchor_paragraph_text"):', result)
except Exception as e:
    print('Error:', e)

# Test setting fields
try:
    img.anchor_paragraph_index = 5
    img.anchor_paragraph_text = 'test alt'
    print('SUCCESS: Set fields - index:', img.anchor_paragraph_index, 'text:', img.anchor_paragraph_text)
except Exception as e:
    print('Error setting fields:', e)

# Test default values
try:
    img2 = Image(image_id='test2', relationship_id='rId2', source_path='word/media/img2.png')
    print('SUCCESS: Default values - index:', img2.anchor_paragraph_index, 'text:', img2.anchor_paragraph_text)
except Exception as e:
    print('Error default values:', e)