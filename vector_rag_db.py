import os
import re
import json
from docx import Document
import chromadb
from chromadb.utils import embedding_functions
import hashlib
import win32com.client
import time
import logging
import traceback
from docx.shared import Pt
import PyPDF2
import textwrap
from collections import deque

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("app.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class VectorRAGDatabase:
    def __init__(self, documents_dir: str, vector_db_path: str):
        """Конструктор класса, принимающий обязательные аргументы"""
        self.documents_dir = documents_dir
        self.vector_db_path = vector_db_path
        
        self.client = chromadb.PersistentClient(path=vector_db_path)
        self.embedding_func = embedding_functions.DefaultEmbeddingFunction()
        
        self.collection = self.client.get_or_create_collection(
            name="documents",
            embedding_function=self.embedding_func
        )
        logger.info(f"Векторная база инициализирована. Путь: {vector_db_path}")

    def convert_doc_to_docx(self, doc_path: str) -> str:
        """Конвертирует .doc в .docx"""
        try:
            logger.info(f"Конвертация: {os.path.basename(doc_path)} в .docx")
            word = win32com.client.Dispatch("Word.Application")
            word.visible = False
            word.DisplayAlerts = False
            
            doc = word.Documents.Open(doc_path)
            new_path = doc_path + "x"
            doc.SaveAs2(new_path, FileFormat=16)
            doc.Close()
            word.Quit()
            
            os.remove(doc_path)
            logger.info(f"Успешно сконвертирован: {os.path.basename(doc_path)} -> {os.path.basename(new_path)}")
            return new_path
        except Exception as e:
            logger.error(f"Ошибка конвертации: {str(e)}")
            return doc_path

    def read_doc(self, file_path: str) -> str:
        """Чтение .doc файла"""
        try:
            logger.info(f"Чтение .doc файла: {os.path.basename(file_path)}")
            word = win32com.client.Dispatch("Word.Application")
            word.visible = False
            word.DisplayAlerts = False
            
            doc = word.Documents.Open(
                FileName=file_path,
                ReadOnly=True,
                ConfirmConversions=False,
                AddToRecentFiles=False
            )
            
            text = doc.Content.Text
            doc.Close()
            word.Quit()
            
            return text
        except Exception as e:
            logger.error(f"Ошибка чтения .doc: {str(e)}")
            return ""

    def read_docx(self, file_path: str) -> str:
        """Чтение .docx файла"""
        try:
            logger.info(f"Чтение .docx файла: {os.path.basename(file_path)}")
            doc = Document(file_path)
            return "\n".join(para.text for para in doc.paragraphs if para.text.strip())
        except Exception as e:
            logger.error(f"Ошибка чтения .docx: {str(e)}")
            return ""

    def read_pdf(self, file_path: str) -> str:
        """Чтение PDF файла"""
        try:
            logger.info(f"Чтение PDF файла: {os.path.basename(file_path)}")
            text = ""
            with open(file_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                for page in reader.pages:
                    text += page.extract_text() + "\n"
            return text
        except Exception as e:
            logger.error(f"Ошибка чтения PDF: {str(e)}")
            return ""

    def chunk_text(self, text: str, chunk_size: int = 1000, overlap: int = 200) -> list:
        """Интеллектуальное разбиение текста на чанки"""
        sentences = re.split(r'(?<=[.!?])\s+', text)
        
        chunks = []
        current_chunk = []
        current_length = 0
        overlap_buffer = deque(maxlen=5) 
        
        for sentence in sentences:
            sentence = sentence.strip()
            if not sentence:
                continue
                
            sentence_length = len(sentence)
            
            if sentence_length > chunk_size:
                wrapped = textwrap.wrap(sentence, width=chunk_size)
                for part in wrapped:
                    sentences.append(part)
                continue
            
            if current_length + sentence_length <= chunk_size:
                current_chunk.append(sentence)
                current_length += sentence_length
                overlap_buffer.append(sentence)
            else:
                chunks.append(" ".join(current_chunk))
                
                current_chunk = list(overlap_buffer)
                current_chunk.append(sentence)
                current_length = sum(len(s) for s in current_chunk)
                overlap_buffer.append(sentence)
        
        if current_chunk:
            chunks.append(" ".join(current_chunk))
        
        return chunks

    def generate_id(self, source: str, chunk_index: int) -> str:
        """Генерация уникального ID"""
        return hashlib.md5(f"{source}_{chunk_index}".encode('utf-8')).hexdigest()


    def update_documents(self, chunk_size: int = 1000, overlap: int = 200):
        """Инкрементное обновление базы (только новые/измененные файлы)"""
        existing_sources = set()
        try:
            results = self.collection.get(include=["metadatas"])
            
            if results and "metadatas" in results and results["metadatas"] is not None:
                for metadata in results["metadatas"]:
                    if metadata and 'source' in metadata:
                        existing_sources.add(metadata['source'])
            else:
                logger.info("Векторная база пуста, будет выполнена полная индексация")
        except Exception as e:
            logger.error(f"Ошибка получения метаданных: {str(e)}")
            logger.error(traceback.format_exc())

        supported_formats = ('.doc', '.docx', '.pdf')
        all_files = os.listdir(self.documents_dir)
        new_files = [
            f for f in all_files
            if (f.lower().endswith(supported_formats) 
                and f not in existing_sources 
                and not f.startswith(('~$',)))
            ]

        if not new_files:
            logger.info("Новых файлов для обработки не найдено")
            return 0, 0

        files_to_process = []
        for filename in new_files:
            if filename.lower().endswith('.doc'):
                file_path = os.path.join(self.documents_dir, filename)
                new_path = self.convert_doc_to_docx(file_path)
                if new_path != file_path:
                    new_filename = os.path.basename(new_path)
                    files_to_process.append(new_filename)
                else:
                    files_to_process.append(filename)
            else:
                files_to_process.append(filename)

        processed_files = 0
        total_chunks = 0
        
        for filename in files_to_process:
            file_path = os.path.join(self.documents_dir, filename)
            logger.info(f"Обработка файла: {filename}")
            
            try:
                start_time = time.time()
                
                if filename.lower().endswith('.doc'):
                    text = self.read_doc(file_path)
                elif filename.lower().endswith('.docx'):
                    text = self.read_docx(file_path)
                elif filename.lower().endswith('.pdf'):
                    text = self.read_pdf(file_path)
                else:
                    continue
                
                if not text.strip():
                    logger.warning(f"Пустой файл: {filename}")
                    continue
                
                chunks = self.chunk_text(text, chunk_size, overlap)
                
                ids = [self.generate_id(filename, i) for i in range(len(chunks))]
                metadatas = [{"source": filename, "chunk_index": i} for i in range(len(chunks))]
                
                self.collection.add(
                    ids=ids,
                    documents=chunks,
                    metadatas=metadatas
                )
                
                processed_files += 1
                total_chunks += len(chunks)
                
                elapsed = time.time() - start_time
                logger.info(f"Добавлено {len(chunks)} чанков за {elapsed:.2f} сек")
                
            except Exception as e:
                logger.error(f"Ошибка обработки файла {filename}: {str(e)}")
                logger.error(traceback.format_exc())
        
        logger.info(f"Обновление завершено. Новых файлов: {processed_files}, Чанков: {total_chunks}")
        return processed_files, total_chunks


    def index_documents(self, chunk_size: int = 1000, overlap: int = 200):
        """Индексация всех документов в директории"""
        processed_files = 0
        total_chunks = 0
        
        supported_formats = ('.doc', '.docx', '.pdf')
        files_to_process = [
            f for f in os.listdir(self.documents_dir) 
            if f.endswith(supported_formats) and not f.startswith(('~$',))
        ]
        
        for filename in files_to_process[:]:
            if filename.endswith('.doc'):
                file_path = os.path.join(self.documents_dir, filename)
                new_path = self.convert_doc_to_docx(file_path)
                if new_path != file_path:
                    new_filename = os.path.basename(new_path)
                    files_to_process[files_to_process.index(filename)] = new_filename
        
        for filename in files_to_process:
            file_path = os.path.join(self.documents_dir, filename)
            logger.info(f"Обработка файла: {filename}")
            
            try:
                start_time = time.time()
                
                if filename.endswith('.doc'):
                    text = self.read_doc(file_path)
                elif filename.endswith('.docx'):
                    text = self.read_docx(file_path)
                elif filename.endswith('.pdf'):
                    text = self.read_pdf(file_path)
                else:
                    continue
                
                if not text.strip():
                    logger.warning(f"Пустой файл: {filename}")
                    continue
                
                chunks = self.chunk_text(text, chunk_size, overlap)
                
                ids = [self.generate_id(filename, i) for i in range(len(chunks))]
                metadatas = [{"source": filename, "chunk_index": i} for i in range(len(chunks))]
                
                self.collection.add(
                    ids=ids,
                    documents=chunks,
                    metadatas=metadatas
                )
                
                processed_files += 1
                total_chunks += len(chunks)
                
                elapsed = time.time() - start_time
                logger.info(f"Добавлено {len(chunks)} чанков за {elapsed:.2f} сек")
                
            except Exception as e:
                logger.error(f"Ошибка обработки файла {filename}: {str(e)}")
                logger.error(traceback.format_exc())
        
        logger.info(f"Индексация завершена. Файлов: {processed_files}, Чанков: {total_chunks}")
        return processed_files, total_chunks

    def search_relevant_chunks(self, query: str, n_results: int = 5) -> list:
        """Поиск релевантных фрагментов"""
        try:
            results = self.collection.query(
                query_texts=[query],
                n_results=n_results
            )
            
            relevant_chunks = []
            for i in range(len(results['ids'][0])):
                relevant_chunks.append({
                    "content": results['documents'][0][i],
                    "source": results["metadatas"][0][i]["source"],
                    "chunk_index": results["metadatas"][0][i]["chunk_index"],
                    "score": results["distances"][0][i]
                })
            
            relevant_chunks.sort(key=lambda x: x["score"], reverse=True)
            
            logger.info(f"Найдено релевантных фрагментов: {len(relevant_chunks)}")
            return relevant_chunks
        except Exception as e:
            logger.error(f"Ошибка поиска: {str(e)}")
            return []
