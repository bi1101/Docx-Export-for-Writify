function extractRawCommentsFromHTML() {
    const vocabElements = document.querySelectorAll(".upgrade_vocab");
    const rawComments = [];

    vocabElements.forEach((element) => {
        const originalVocab = element.querySelector(".original-vocab").innerText;
        const improvedVocab = element.querySelector(".improved-vocab").innerText;
        const explanation = element.querySelector(".explanation").innerText;

        const commentData = {
            originalVocab: originalVocab,
            improvedVocab: improvedVocab,
            explanation: explanation
        };

        rawComments.push(commentData);
    });

    return rawComments;
}

function convertRawCommentsToDocxFormat(rawComments) {
    return rawComments.map((comment, index) => ({
        id: index,
        author: "Teacher",
        date: new Date(),
        children: [
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: comment.originalVocab + " -> " + comment.improvedVocab
                    })
                ]
            }),
            new docx.Paragraph({}),
            new docx.Paragraph({
                children: [new docx.TextRun({ text: comment.explanation })]
            })
        ]
    }));
}

function createSectionsWithComments(rawComments) {
    const essayText = document.querySelector("#my-text").innerText;
    const essayPrompt = document
        .querySelector(".essay_prompt .elementor-widget-container")
        .innerText.trim();
    const essayParagraphs = essayText.split(/\\r?\\n/).map((p) => p.trimStart());
    const essayPromptParagraphs = essayPrompt
        .split(/\\r?\\n/)
        .map((p) => p.trimStart());
    const outputParagraphs = [];

    // Add the essay prompt paragraphs to the output
    for (let promptParagraph of essayPromptParagraphs) {
        if (promptParagraph.trim()) {
            // Check if paragraph is not just whitespace
            outputParagraphs.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: promptParagraph,
                            bold: true
                        })
                    ]
                })
            );
        }
    }

    for (let paraText of essayParagraphs) {
        if (paraText.trim()) {
            // Check if paragraph is not just whitespace
            let currentPosition = 0;
            const paraChildren = [];
            let localCommentIndex = 0; // Reset for each paragraph

            while (localCommentIndex < rawComments.length) {
                const commentStartPos = paraText
                    .toLowerCase()
                    .indexOf(
                        rawComments[localCommentIndex].originalVocab.toLowerCase(),
                        currentPosition
                    );

                if (commentStartPos !== -1) {
                    // Add text before the comment
                    const preCommentText = paraText.slice(
                        currentPosition,
                        commentStartPos
                    );
                    paraChildren.push(new docx.TextRun(preCommentText));

                    // Add the comment
                    paraChildren.push(new docx.CommentRangeStart(localCommentIndex));
                    paraChildren.push(
                        new docx.TextRun(rawComments[localCommentIndex].originalVocab)
                    );
                    paraChildren.push(new docx.CommentRangeEnd(localCommentIndex));
                    paraChildren.push(
                        new docx.TextRun({
                            children: [new docx.CommentReference(localCommentIndex)]
                        })
                    );

                    currentPosition =
                        commentStartPos +
                        rawComments[localCommentIndex].originalVocab.length;

                    localCommentIndex++;
                } else {
                    console.warn(
                        `Skipped raw comment at index ${localCommentIndex} because it was not found in the essay text.`
                    );
                    // If no comment is found in the current paragraph, move on to the next comment
                    localCommentIndex++;
                    continue;
                }
            }

            // Add the remaining part of the paragraph
            const postCommentText = paraText.slice(currentPosition);
            paraChildren.push(new docx.TextRun(postCommentText));

            outputParagraphs.push(new docx.Paragraph({ children: paraChildren }));
        }
    }

    return outputParagraphs;
}

function exportDocument() {
    const rawComments = extractRawCommentsFromHTML();
    const commentsForDocx = convertRawCommentsToDocxFormat(rawComments);

    // Generating sections
    const sectionsChildren = [];
    sectionsChildren.push(...createSectionsWithComments(rawComments));
    sectionsChildren.push(...createNormalSections("task-response"));
    /*sectionsChildren.push(...createNormalSections("coherence-cohesion"));
    sectionsChildren.push(...createNormalSections("lexical-resource"));
    sectionsChildren.push(...createNormalSections("grammatical-range-accuracy"));
    sectionsChildren.push(...createNormalSections("sample-answer"));
  */
    console.dir(sectionsChildren);
    const doc = new docx.Document({
        comments: {
            children: commentsForDocx
        },
        sections: [
            {
                properties: {},
                children: sectionsChildren
            }
        ]
    });

    // Convert the document to a blob and save it
    docx.Packer.toBlob(doc).then((blob) => {
        saveBlobAsDocx(blob);
    });
}

function createNormalSections(className) {
    const element = document.querySelector(
        `.${className} .elementor-widget-container .elementor-shortcode`
    );
    if (!element) {
        console.warn(`No element found with class name: ${className}`);
        return [];
    }

    const sections = [];
    element.childNodes.forEach((child) => {
        if (child.nodeType === 1) {
            // Check if the node is an element
            if (child.tagName === "P") {
                // For paragraph tags
                sections.push(htmlParagraphToDocx(child.outerHTML));
            } else if (child.tagName === "OL" || child.tagName === "UL") {
                // For ordered or unordered lists
                sections.push(bulletPointsToDocx(child.outerHTML));
            }
        }
    });

    return sections;
}

function htmlParagraphToDocx(htmlContent) {
    // Convert the HTML string into a DOM element
    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = htmlContent;

    const paragraph = tempDiv.querySelector("p");
    if (!paragraph) {
        console.warn("No paragraph element found in the provided HTML content.");
        return;
    }

    // Use processNodeForFormatting to handle child nodes
    const children = [];
    Array.from(paragraph.childNodes).forEach((child) => {
        children.push(...processNodeForFormatting(child));
    });

    return new docx.Paragraph({ children });
}

function processNodeForFormatting(node) {
    let textRuns = [];

    // Handle text nodes
    if (node.nodeType === 3) {
        // Node type 3 is a Text node
        textRuns.push(new docx.TextRun(node.nodeValue));
    }

    // Handle element nodes like <strong>, <em>, etc.
    else if (node.nodeType === 1) {
        // Node type 1 is an Element node
        const textContent = node.innerText;

        // Basic formatting options
        let formattingOptions = {};

        // Check the tag to determine formatting
        switch (node.tagName) {
            case "STRONG":
            case "B":
                formattingOptions.bold = true;
                break;
            case "EM":
            case "I":
                formattingOptions.italic = true;
                break;
            case "U":
                formattingOptions.underline = {
                    color: "auto",
                    type: docx.UnderlineType.SINGLE
                };
                break;
            // Add cases for other formatting tags as needed
        }

        // Check for nested formatting
        if (node.children.length > 0) {
            Array.from(node.childNodes).forEach((childNode) => {
                textRuns.push(...processNodeForFormatting(childNode));
            });
        } else {
            textRuns.push(
                new docx.TextRun({
                    text: textContent,
                    ...formattingOptions
                })
            );
        }
    }

    return textRuns;
}

function bulletPointsToDocx(outerHTML) {
    console.log("Processing bullet point content:", outerHTML);
    // You can later replace the above log statement with the actual logic
}

function saveBlobAsDocx(blob) {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "ielts-comments.docx";
    document.body.appendChild(a);
    a.click();
    a.remove();
}

// Event Listener
document.addEventListener("DOMContentLoaded", function () {
    document.getElementById("export-docx").addEventListener("click", function () {
        exportDocument();
    });
});
