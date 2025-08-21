"use client";

import { useState, useEffect, useRef } from "react";
import { Button } from "@/components/ui/button";
import { Copy, Check } from "lucide-react";
import { cn } from "@/lib/utils";

interface CodeBlockProps {
  code: string;
  language?: string;
  title?: string;
  description?: string;
  className?: string;
}

export function CodeBlock({
  code,
  language = "vba",
  title,
  description,
  className,
}: CodeBlockProps) {
  const [copied, setCopied] = useState(false);
  const codeRef = useRef<HTMLElement>(null);

  const copyToClipboard = async () => {
    try {
      await navigator.clipboard.writeText(code);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.error("Error al copiar:", err);
    }
  };

  useEffect(() => {
    if (codeRef.current && language === "vba") {
      const element = codeRef.current;
      element.innerHTML = code;

      let html = element.innerHTML;

      // Comentarios (solo comilla simple)
      html = html.replace(
        /^(\s*)'(.*)$/gm,
        '<span class="vba-comment">\'$2</span>'
      );

      // Keywords
      const keywords = [
        "Sub",
        "End Sub",
        "Function",
        "End Function",
        "Dim",
        "As",
        "Set",
        "If",
        "Then",
        "Else",
        "ElseIf",
        "End If",
        "For",
        "To",
        "Next",
        "Step",
        "While",
        "Wend",
        "Do",
        "Loop",
        "Until",
        "Exit",
        "Call",
        "Return",
        "Public",
        "Private",
        "Static",
        "Const",
        "Optional",
        "ByRef",
        "ByVal",
        "With",
        "End With",
        "Select",
        "Case",
        "End Select",
        "On Error",
        "Resume",
        "GoTo",
        "ReDim",
        "Preserve",
      ];

      keywords.forEach((keyword) => {
        const regex = new RegExp(`\\b${keyword}\\b`, "gi");
        html = html.replace(
          regex,
          `<span class="vba-keyword">${keyword}</span>`
        );
      });

      // Tipos de datos
      const types = [
        "String",
        "Integer",
        "Long",
        "Double",
        "Single",
        "Boolean",
        "Date",
        "Variant",
        "Object",
        "Worksheet",
        "Workbook",
        "Range",
        "Application",
        "Collection",
      ];

      types.forEach((type) => {
        const regex = new RegExp(`\\b${type}\\b`, "gi");
        html = html.replace(regex, `<span class="vba-type">${type}</span>`);
      });

      // Números
      html = html.replace(
        /\b\d+(\.\d+)?\b/g,
        '<span class="vba-number">$&</span>'
      );

      element.innerHTML = html;
    }
  }, [code, language]);

  return (
    <div
      className={cn(
        "border rounded-lg overflow-hidden bg-card w-full max-w-full",
        className
      )}
    >
      {title && (
        <div className="bg-muted/50 px-4 py-2 border-b">
          <div className="flex items-center justify-between">
            <h4 className="font-semibold text-sm">{title}</h4>
            <Button
              variant="ghost"
              size="sm"
              onClick={copyToClipboard}
              className="h-8 px-2"
            >
              {copied ? (
                <Check className="h-4 w-4 text-green-600" />
              ) : (
                <Copy className="h-4 w-4" />
              )}
              <span className="ml-1 text-xs">
                {copied ? "¡Copiado!" : "Copiar"}
              </span>
            </Button>
          </div>
          {description && (
            <p className="text-xs text-muted-foreground mt-1">{description}</p>
          )}
        </div>
      )}
      <div className="relative w-full max-w-full overflow-x-auto overflow-y-auto">
        <div className="w-full max-w-full">
          <pre className="vba-code-block whitespace-pre  md:w-full w-[80vw]">
            <code ref={codeRef} className={`language-${language} block w-full`}>
              {code}
            </code>
          </pre>
        </div>

        {!title && (
          <Button
            variant="ghost"
            size="sm"
            onClick={copyToClipboard}
            className="absolute top-2 right-2 h-8 px-2 bg-gray-800/80 hover:bg-gray-700/80 text-gray-200"
          >
            {copied ? (
              <Check className="h-4 w-4 text-green-400" />
            ) : (
              <Copy className="h-4 w-4" />
            )}
          </Button>
        )}
      </div>
    </div>
  );
}
