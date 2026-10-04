import React, { useState } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { VideoTutorialCard } from './VideoTutorialCard';

interface VideoTutorialModalProps {
  isOpen: boolean;
  onClose: () => void;
  setView?: (view: any) => void;
}

export const VideoTutorialModal: React.FC<VideoTutorialModalProps> = ({
  isOpen,
  onClose,
  setView
}) => {
  return (
    <AnimatePresence>
      {isOpen && (
        <div 
          className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-6 bg-slate-950/80 backdrop-blur-md overflow-y-auto"
          onClick={onClose}
        >
          <motion.div
            initial={{ opacity: 0, scale: 0.95, y: 15 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.95, y: 15 }}
            transition={{ duration: 0.2 }}
            className="w-full max-w-6xl my-auto relative"
            onClick={(e) => e.stopPropagation()}
          >
            <VideoTutorialCard 
              setView={setView} 
              onClose={onClose} 
              isModal={true} 
            />
          </motion.div>
        </div>
      )}
    </AnimatePresence>
  );
};
