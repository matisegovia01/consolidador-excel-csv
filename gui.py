from app.shell import ConsolidatorApp
import sys


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "--self-test":
        from app.diagnostics import verificar
        verificar(sys.argv[2])
    else:
        ConsolidatorApp().mainloop()
