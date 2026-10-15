from oa.server import main
if __name__=='__main__':
    import sys
    if len(sys.argv)>2 and sys.argv[1]=='--diagnose':
        from oa.core import extract
        import json,traceback
        try:
            text,method,warnings=extract(sys.argv[2]);print(json.dumps({'characters':len(text),'method':method,'warnings':warnings},ensure_ascii=False))
        except Exception:
            traceback.print_exc();sys.exit(1)
    else:main()
