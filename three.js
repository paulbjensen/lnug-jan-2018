// Dependencies
const fs = require('fs');
const officegen = require('officegen');

// Instantiate a PowerPoint file
const pptx = officegen('pptx');

// Set the document title
pptx.setDocTitle('My Powerpoint deck');

// Set the size of the slides
pptx.setSlideSize('screen16x9');

// Create the first slide
const slide = pptx.makeNewSlide();

// Set the name of the slide
slide.name = 'Title slide';

// Set the background and foreground colours
slide.back = '492c14';
slide.color = 'e7dd85';

/* Add some text to the slide */
slide.addText('Hello LNUG', {
	font_face: 'Helvetica Neue',
	font_size: 48,
	align: 'center',
	x: '25%',
	y: '45%',
	cx: '50%',
	cy: '10%'
});

// Create a writeable file stream
const out = fs.createWriteStream('example.pptx');

// Bind a log message when the file stream is closed
const closed = () => console.log('Finished to create the PPTX file!');
out.on('close', closed);

// Pipe the pptx file generation stream to the writeable file stream
pptx.generate(out);
