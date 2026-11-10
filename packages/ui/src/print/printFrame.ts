/**
 * Print a document from a plain browser: load it in an off-screen frame and open the browser's
 * print dialog for it (which also offers Save as PDF). Resolves once the dialog has closed.
 */
export function printInFrame(html: string): Promise<void> {
  return new Promise((resolve) => {
    const frame = document.createElement('iframe');
    frame.setAttribute('aria-hidden', 'true');
    frame.tabIndex = -1;
    // Off screen rather than hidden: some browsers print nothing from a hidden frame.
    Object.assign(frame.style, {
      position: 'fixed',
      left: '-10000px',
      top: '0',
      width: '800px',
      height: '600px',
      border: '0',
    });
    let done = false;
    const finish = () => {
      if (done) return;
      done = true;
      frame.remove();
      resolve();
    };
    frame.onload = () => {
      const win = frame.contentWindow;
      if (!win) return finish();
      win.addEventListener('afterprint', () => setTimeout(finish, 0));
      win.focus();
      // print() blocks until the dialog closes in most browsers; afterprint covers the rest.
      win.print();
      setTimeout(finish, 60_000);
    };
    frame.srcdoc = html;
    document.body.append(frame);
  });
}
