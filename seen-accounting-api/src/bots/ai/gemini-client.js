import { env } from '../../config/env.js';
import { logger } from '../../shared/logger.js';

const GEMINI_BASE_URL = 'https://generativelanguage.googleapis.com/v1beta/models';

const FALLBACK_MODELS = [
  'gemini-2.0-flash',
  'gemini-2.0-flash-001',
  'gemini-flash-latest',
  'gemini-pro-latest',
  'gemini-1.5-flash',
  'gemini-pro',
];

const GENERATION_CONFIG = {
  temperature: 0.1,
  topP: 0.8,
  topK: 40,
  maxOutputTokens: 2048,
};

let workingModel = null;

/**
 * Call Gemini API with automatic model fallback
 */
export async function callGemini(systemPrompt, userMessage) {
  const apiKey = env.geminiApiKey;
  if (!apiKey) throw new Error('GEMINI_API_KEY not configured');

  const models = workingModel
    ? [workingModel, ...FALLBACK_MODELS.filter((m) => m !== workingModel)]
    : FALLBACK_MODELS;

  for (const model of models) {
    try {
      const url = `${GEMINI_BASE_URL}/${model}:generateContent?key=${apiKey}`;

      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          systemInstruction: { parts: [{ text: systemPrompt }] },
          contents: [{ parts: [{ text: userMessage }] }],
          generationConfig: GENERATION_CONFIG,
          safetySettings: [
            { category: 'HARM_CATEGORY_HARASSMENT', threshold: 'BLOCK_NONE' },
            { category: 'HARM_CATEGORY_HATE_SPEECH', threshold: 'BLOCK_NONE' },
            { category: 'HARM_CATEGORY_SEXUALLY_EXPLICIT', threshold: 'BLOCK_NONE' },
            { category: 'HARM_CATEGORY_DANGEROUS_CONTENT', threshold: 'BLOCK_NONE' },
          ],
        }),
      });

      if (!response.ok) {
        const errorText = await response.text();
        logger.warn(`Gemini model ${model} failed: ${response.status} - ${errorText}`);
        continue;
      }

      const data = await response.json();
      const text = data.candidates?.[0]?.content?.parts?.[0]?.text;

      if (!text) {
        logger.warn(`Gemini model ${model} returned empty response`);
        continue;
      }

      // Cache working model
      workingModel = model;
      return text;
    } catch (err) {
      logger.warn(`Gemini model ${model} error: ${err.message}`);
      continue;
    }
  }

  throw new Error('All Gemini models failed');
}

/**
 * Parse JSON from Gemini response (handles markdown code blocks)
 */
export function parseGeminiJSON(text) {
  // Remove markdown code blocks
  let cleaned = text.replace(/```json\s*/g, '').replace(/```\s*/g, '');

  // Try to find JSON object
  const jsonMatch = cleaned.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    throw new Error('No JSON found in Gemini response');
  }

  return JSON.parse(jsonMatch[0]);
}
